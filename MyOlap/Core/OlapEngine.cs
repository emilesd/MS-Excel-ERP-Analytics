using System.Text.RegularExpressions;
using MyOlap.Data;

namespace MyOlap.Core;

/// <summary>
/// Central engine managing the active model, current view state,
/// undo history, and all OLAP operations (drill, swap, keep, remove).
/// </summary>
public class OlapEngine
{
    private static readonly Lazy<OlapEngine> _lazy = new(() => new OlapEngine());
    public static OlapEngine Instance => _lazy.Value;

    private readonly SqliteRepository _repo = SqliteRepository.Instance;
    private readonly UndoManager _undo = new();

    public ViewState? CurrentView { get; private set; }
    public OlapModel? ActiveModel { get; private set; }
    private readonly Dictionary<string, ViewState> _sheetViews = new();

    // Persistent caches — survive multiple BuildGrid calls, cleared only when data actually changes.
    private long _cachedFactsModelId = -1;
    private List<(FactData fact, Dictionary<long, long> members)>? _cachedParsedFacts;
    private readonly Dictionary<long, Dictionary<long, int>> _persistentDescendantCache = new();

    // Clears only the parsed-fact cache. Call after data load/clear — member hierarchy is unchanged.
    public void InvalidateFactCache()
    {
        _cachedFactsModelId = -1;
        _cachedParsedFacts = null;
    }

    // Full reset: facts + descendant sign maps. Used when switching to a different model.
    private void InvalidateModelCache()
    {
        InvalidateFactCache();
        _persistentDescendantCache.Clear();
    }

    public void SaveViewForSheet(string sheetName)
    {
        if (CurrentView != null && !string.IsNullOrEmpty(sheetName))
            _sheetViews[sheetName] = CurrentView.Clone();
    }

    public bool RestoreViewForSheet(string sheetName)
    {
        if (!string.IsNullOrEmpty(sheetName) && _sheetViews.TryGetValue(sheetName, out var view))
        {
            CurrentView = view;
            return true;
        }
        return false;
    }

    // Model-guarded restore: only adopts the saved view when it belongs to the given model,
    // so a sheet reused across models can never resurrect a foreign layout.
    public bool RestoreViewForSheet(string sheetName, long expectedModelId)
    {
        if (!string.IsNullOrEmpty(sheetName) && _sheetViews.TryGetValue(sheetName, out var view)
            && view.ModelId == expectedModelId)
        {
            CurrentView = view;
            return true;
        }
        return false;
    }

    // Adopts a view reconstructed from the worksheet's header metadata (same-model reselect).
    // Refuses views belonging to a different model.
    public void SetCurrentView(ViewState view)
    {
        if (view != null && ActiveModel != null && view.ModelId == ActiveModel.Id)
            CurrentView = view;
    }

    /// <summary>
    /// Sets ActiveModel without rebuilding the default view — used when reopening a workbook
    /// that already has a report layout so the Info label and ribbon commands reconnect.
    /// </summary>
    public bool ActivateModel(long modelId)
    {
        if (ActiveModel?.Id == modelId) return true;
        var model = _repo.GetAllModels().FirstOrDefault(m => m.Id == modelId);
        if (model == null) return false;
        InvalidateModelCache();
        ActiveModel = model;
        CurrentView = null; // caller restores layout from sheet/meta/DB
        return true;
    }

    /// <summary>
    /// Persists the current layout to SQLite (same store as dimensions/settings),
    /// so re-selecting the model after Excel restart restores the last used view.
    /// </summary>
    public void PersistCurrentView()
    {
        if (CurrentView == null || ActiveModel == null) return;
        if (CurrentView.ModelId != ActiveModel.Id) return;
        if (CurrentView.RowAxes.Count == 0 || CurrentView.ColAxes.Count == 0) return;
        _repo.SaveModelView(ActiveModel.Id, ViewStateCodec.Serialize(CurrentView));
    }

    /// <summary>
    /// Loads the last persisted layout for a model into CurrentView.
    /// Returns false when none is stored or the payload is invalid.
    /// </summary>
    public bool TryLoadPersistedView(long modelId)
    {
        var payload = _repo.LoadModelView(modelId);
        var view = ViewStateCodec.Deserialize(payload, id =>
            _repo.GetDimensions(modelId).FirstOrDefault(d => d.Id == id)?.Name);
        if (view == null || view.ModelId != modelId) return false;
        if (ActiveModel == null || ActiveModel.Id != modelId) return false;
        CurrentView = view;
        return true;
    }

    public void ClearPersistedView(long modelId) => _repo.DeleteModelView(modelId);

    /// <summary>
    /// Keeps the current row/column/POV layout after Manage Model changes, but:
    /// - drops dimensions that were deleted
    /// - places any brand-new dimension (not yet on an axis) onto POV with its default root
    /// Does NOT reset drilled members or axis arrangement to the model default view.
    /// </summary>
    public void SyncViewWithModelStructure()
    {
        if (CurrentView == null || ActiveModel == null) return;
        var dims = _repo.GetDimensions(ActiveModel.Id);
        var dimIds = dims.Select(d => d.Id).ToHashSet();

        CurrentView.RowAxes.RemoveAll(a => !dimIds.Contains(a.DimensionId));
        CurrentView.ColAxes.RemoveAll(a => !dimIds.Contains(a.DimensionId));
        foreach (var orphan in CurrentView.PovSelections.Keys.Where(k => !dimIds.Contains(k)).ToList())
            CurrentView.PovSelections.Remove(orphan);

        EnsureAxesNonEmpty();

        var placed = new HashSet<long>(
            CurrentView.RowAxes.Select(a => a.DimensionId)
                .Concat(CurrentView.ColAxes.Select(a => a.DimensionId))
                .Concat(CurrentView.PovSelections.Keys));

        foreach (var d in dims.OrderBy(d => d.SortOrder))
        {
            if (placed.Contains(d.Id)) continue;
            var roots = _repo.GetRootMembers(d.Id);
            if (roots.Count == 0) continue;
            var best = roots.FirstOrDefault(r => _repo.GetChildren(r.Id).Count > 0) ?? roots[0];
            CurrentView.PovSelections[d.Id] = best.Id;
        }

        // Member hierarchy may have changed (load/add/remove); keep fact cache.
        _persistentDescendantCache.Clear();
    }

    /// <summary>
    /// Rebuilds a ViewState from worksheet header metadata cells ("DIM:x|MBR:y" per header cell),
    /// so re-selecting the same model refreshes the data while keeping the layout currently on
    /// the sheet instead of resetting to the default view.
    /// Classification: cell (1,1) present → row 1 is the POV row; a row &gt; 1 containing a
    /// metadata cell in column 1 is a row-header row (one axis per column); any other row is a
    /// column-header row (one axis per row). Returns null when no intact grid is found.
    /// </summary>
    public ViewState? BuildViewFromMetadata(long modelId, IReadOnlyList<(int row, int col, long dimId, long mbrId)> cells)
    {
        var dims = _repo.GetDimensions(modelId);
        if (dims.Count == 0 || cells.Count == 0) return null;
        var dimMap = dims.ToDictionary(d => d.Id);

        bool hasPovRow = cells.Any(c => c.row == 1 && c.col == 1);
        var rowsWithCol1 = new HashSet<int>(cells.Where(c => c.col == 1 && c.row > 1).Select(c => c.row));
        var pov = new Dictionary<long, long>();
        var rowAxisCells = new SortedDictionary<int, List<(int row, long dimId, long mbrId)>>();
        var colAxisCells = new SortedDictionary<int, List<(int col, long dimId, long mbrId)>>();

        foreach (var (r, c, dimId, mbrId) in cells)
        {
            if (!dimMap.ContainsKey(dimId)) continue;
            var member = _repo.GetMember(mbrId);
            if (member == null || member.DimensionId != dimId) continue;

            if (r == 1 && hasPovRow) { pov[dimId] = mbrId; continue; }
            if (rowsWithCol1.Contains(r))
            {
                if (!rowAxisCells.TryGetValue(c, out var rl)) rowAxisCells[c] = rl = new();
                rl.Add((r, dimId, mbrId));
            }
            else
            {
                if (!colAxisCells.TryGetValue(r, out var cl)) colAxisCells[r] = cl = new();
                cl.Add((c, dimId, mbrId));
            }
        }

        var view = new ViewState { ModelId = modelId };
        foreach (var (dimId, memberId) in pov)
            view.PovSelections[dimId] = memberId;

        foreach (var (_, list) in rowAxisCells) // columns left-to-right = axis order
        {
            if (list.Any(x => x.dimId != list[0].dimId)) return null; // mixed dims → damaged metadata
            var ids = new List<long>();
            foreach (var item in list.OrderBy(x => x.row))
                if (!ids.Contains(item.mbrId)) ids.Add(item.mbrId);
            if (ids.Count > 0)
                view.RowAxes.Add(new DimensionAxis { DimensionId = list[0].dimId, DimensionName = dimMap[list[0].dimId].Name, VisibleMemberIds = ids });
        }
        foreach (var (_, list) in colAxisCells) // rows top-to-bottom = axis order
        {
            if (list.Any(x => x.dimId != list[0].dimId)) return null; // mixed dims → damaged metadata
            var ids = new List<long>();
            foreach (var item in list.OrderBy(x => x.col))
                if (!ids.Contains(item.mbrId)) ids.Add(item.mbrId);
            if (ids.Count > 0)
                view.ColAxes.Add(new DimensionAxis { DimensionId = list[0].dimId, DimensionName = dimMap[list[0].dimId].Name, VisibleMemberIds = ids });
        }

        // An intact MyOlap grid always has both axes populated (EnsureAxesNonEmpty invariant).
        if (view.RowAxes.Count == 0 || view.ColAxes.Count == 0) return null;
        return view;
    }

    /// <summary>
    /// Opens a model and builds the default view:
    /// Measures on rows, Time on columns, all other dimensions on POV (page filter).
    /// </summary>
    public ViewState SelectModel(long modelId, bool preserveUndo = false)
    {
        // Capture OLD model id BEFORE updating ActiveModel so modelChanging is correct.
        bool modelChanging = ActiveModel?.Id != modelId;

        var models = _repo.GetAllModels();
        ActiveModel = models.FirstOrDefault(m => m.Id == modelId);
        if (ActiveModel == null)
            throw new InvalidOperationException("Model not found.");

        // Only clear undo when switching to a different model or when explicitly requested.
        // Refreshing the same model (Refresh Data) passes preserveUndo=true so the
        // user's undo history survives the data reload.
        if (!preserveUndo)
            _undo.Clear();
        if (modelChanging)
            InvalidateModelCache();
        else
            _persistentDescendantCache.Clear(); // Member structure may have changed (dimension reload); keep fact cache.

        var dims = _repo.GetDimensions(modelId);

        var view = new ViewState { ModelId = modelId };
        var measureDim = dims.FirstOrDefault(d => d.DimType == DimensionType.Measure);
        var timeDim = dims.FirstOrDefault(d => d.DimType == DimensionType.Time);

        if (measureDim != null)
        {
            var roots = _repo.GetRootMembers(measureDim.Id);
            if (roots.Count > 0)
            {
                view.RowAxes.Add(new DimensionAxis
                {
                    DimensionId = measureDim.Id,
                    DimensionName = measureDim.Name,
                    VisibleMemberIds = roots.Select(r => r.Id).ToList()
                });
            }
        }

        if (timeDim != null)
        {
            var roots = _repo.GetRootMembers(timeDim.Id);
            if (roots.Count > 0)
            {
                view.ColAxes.Add(new DimensionAxis
                {
                    DimensionId = timeDim.Id,
                    DimensionName = timeDim.Name,
                    VisibleMemberIds = roots.Select(r => r.Id).ToList()
                });
            }
        }

        foreach (var d in dims)
        {
            if (d.Id == measureDim?.Id || d.Id == timeDim?.Id) continue;
            var roots = _repo.GetRootMembers(d.Id);
            if (roots.Count > 0)
            {
                var best = roots.FirstOrDefault(r => _repo.GetChildren(r.Id).Count > 0) ?? roots[0];
                view.PovSelections[d.Id] = best.Id;
            }
        }

        CurrentView = view;
        return view;
    }

    /// <summary>
    /// Builds a 2D grid of values for the current view and returns
    /// the row headers, column headers, and value matrix.
    /// </summary>
    public GridResult BuildGrid()
    {
        if (CurrentView == null || ActiveModel == null)
            throw new InvalidOperationException("No model selected.");

        var settings = _repo.GetSettings(CurrentView.ModelId);
        var result = new GridResult();

        var rowMembers = new List<List<Member>>();
        foreach (var axis in CurrentView.RowAxes)
        {
            var members = axis.VisibleMemberIds
                .Select(id => _repo.GetMember(id))
                .Where(m => m != null)
                .Cast<Member>()
                .ToList();
            rowMembers.Add(members);
        }

        var colMembers = new List<List<Member>>();
        foreach (var axis in CurrentView.ColAxes)
        {
            var members = axis.VisibleMemberIds
                .Select(id => _repo.GetMember(id))
                .Where(m => m != null)
                .Cast<Member>()
                .ToList();
            colMembers.Add(members);
        }

        // Row member order is taken verbatim from VisibleMemberIds so that explicit
        // user ordering from Pick Members is always respected.

        var rowCombos = CartesianProduct(rowMembers);
        var colCombos = CartesianProduct(colMembers);

        result.RowHeaders = rowCombos;
        result.ColHeaders = colCombos;
        result.RowDimensionNames = CurrentView.RowAxes.Select(a => a.DimensionName).ToList();
        result.ColDimensionNames = CurrentView.ColAxes.Select(a => a.DimensionName).ToList();

        var dims = _repo.GetDimensions(CurrentView.ModelId);
        var dimOrder = dims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();
        result.Values = new decimal?[rowCombos.Count, colCombos.Count];

        bool factCacheHit = _cachedParsedFacts != null && _cachedFactsModelId == CurrentView.ModelId;
        if (!factCacheHit)
        {
            var allFacts = _repo.GetAllFacts(CurrentView.ModelId);

            _cachedParsedFacts = new List<(FactData fact, Dictionary<long, long> members)>(allFacts.Count);
            foreach (var f in allFacts)
            {
                var parts = f.MemberKey.Split('|');
                var memberMap = new Dictionary<long, long>();
                for (int i = 0; i < dimOrder.Count && i < parts.Length; i++)
                {
                    if (long.TryParse(parts[i], out var mid) && mid > 0)
                        memberMap[dimOrder[i]] = mid;
                }
                _cachedParsedFacts.Add((f, memberMap));
            }
            _cachedFactsModelId = CurrentView.ModelId;
        }
        var parsedFacts = _cachedParsedFacts;

        var descendantCache = _persistentDescendantCache;
        void EnsureDescendants(long memberId)
        {
            if (descendantCache.ContainsKey(memberId)) return;
            var signs = new Dictionary<long, int>();
            signs[memberId] = 1;
            void BuildSigns(long parentId, int parentSign)
            {
                var children = _repo.GetChildren(parentId);
                foreach (var child in children)
                {
                    int childSign;
                    if (child.ConsolOperator == "x") childSign = 0;
                    else if (child.ConsolOperator == "-") childSign = parentSign * -1;
                    else childSign = parentSign;

                    long effectiveId = child.SharedFromId ?? child.Id;
                    if (!signs.ContainsKey(effectiveId))
                        signs[effectiveId] = childSign;
                    signs[child.Id] = childSign;
                    BuildSigns(child.Id, childSign);
                }
            }

            // If this member is a shared copy, include the original and its full sub-hierarchy
            // so fact lookups match regardless of how deep the original's tree goes.
            var rootMem = _repo.GetMember(memberId);
            if (rootMem?.SharedFromId != null)
            {
                signs.TryAdd(rootMem.SharedFromId.Value, 1);
                BuildSigns(rootMem.SharedFromId.Value, 1);
            }

            BuildSigns(memberId, 1);
            descendantCache[memberId] = signs;
        }

        foreach (var mid in CurrentView.PovSelections.Values)
            EnsureDescendants(mid);
        foreach (var combo in rowCombos)
            foreach (var m in combo)
                EnsureDescendants(m.Id);
        foreach (var combo in colCombos)
            foreach (var m in combo)
                EnsureDescendants(m.Id);

        for (int r = 0; r < rowCombos.Count; r++)
        {
            for (int c = 0; c < colCombos.Count; c++)
            {
                var cellMembers = new Dictionary<long, long>(CurrentView.PovSelections);

                foreach (var (axis, idx) in CurrentView.RowAxes.Select((a, i) => (a, i)))
                    cellMembers[axis.DimensionId] = rowCombos[r][idx].Id;

                foreach (var (axis, idx) in CurrentView.ColAxes.Select((a, i) => (a, i)))
                    cellMembers[axis.DimensionId] = colCombos[c][idx].Id;

                decimal total = 0;
                bool anyValue = false;

                foreach (var (fact, factMembers) in parsedFacts)
                {
                    if (!fact.NumericValue.HasValue) continue;

                    bool match = true;
                    int netSign = 1;
                    foreach (var (dimId, targetMemberId) in cellMembers)
                    {
                        if (!factMembers.TryGetValue(dimId, out var factMemberId))
                            continue;
                        var signMap = descendantCache[targetMemberId];
                        if (!signMap.TryGetValue(factMemberId, out var sign))
                        { match = false; break; }
                        if (sign == 0) { match = false; break; }
                        netSign *= sign;
                    }
                    if (match)
                    {
                        total += fact.NumericValue.Value * netSign;
                        anyValue = true;
                    }
                }

                result.Values[r, c] = anyValue ? total : null;
            }
        }

        ApplyTimeBalance(result, rowCombos, colCombos, dims, parsedFacts, descendantCache);
        ApplyFormulas(result, rowCombos, colCombos, dims, parsedFacts, descendantCache);

        if (settings.OmitEmptyRows)
            result.OmitEmptyRows();
        if (settings.OmitEmptyColumns)
            result.OmitEmptyColumns();

        result.MemberDisplay = settings.MemberDisplay;
        return result;
    }

    private void ApplyFormulas(GridResult result, List<List<Member>> rowCombos, List<List<Member>> colCombos,
        List<Dimension> dims,
        List<(FactData fact, Dictionary<long, long> members)> parsedFacts,
        Dictionary<long, Dictionary<long, int>> descendantCache)
    {
        for (int r = 0; r < rowCombos.Count; r++)
        {
            for (int c = 0; c < colCombos.Count; c++)
            {
                var allMembers = rowCombos[r].Concat(colCombos[c]).ToList();
                var formulaMember = allMembers.FirstOrDefault(m => !string.IsNullOrEmpty(m.Formula));

                // When the formula member lives in the header (POV) it won't appear in
                // rowCombos/colCombos, so the standard sibling-scan is skipped entirely.
                // Detect this case and fall through to FillOffGridSiblings only.
                bool formulaMemberInPov = false;
                long formulaDimId = 0;

                if (formulaMember != null)
                {
                    bool isOnRow = rowCombos[r].Contains(formulaMember);
                    int dimIdx = isOnRow
                        ? rowCombos[r].IndexOf(formulaMember)
                        : colCombos[c].IndexOf(formulaMember);
                    formulaDimId = isOnRow
                        ? CurrentView!.RowAxes[dimIdx].DimensionId
                        : CurrentView!.ColAxes[dimIdx].DimensionId;

                    var siblingValues = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);

                    if (isOnRow)
                    {
                        for (int sr = 0; sr < rowCombos.Count; sr++)
                        {
                            if (sr == r) continue;
                            if (dimIdx >= rowCombos[sr].Count) continue;
                            bool sameContext = true;
                            for (int d = 0; d < rowCombos[sr].Count; d++)
                            {
                                if (d == dimIdx) continue;
                                if (d >= rowCombos[r].Count || rowCombos[sr][d].Id != rowCombos[r][d].Id)
                                { sameContext = false; break; }
                            }
                            if (sameContext)
                                siblingValues[rowCombos[sr][dimIdx].Name] = result.Values[sr, c];
                        }
                    }
                    else
                    {
                        for (int sc = 0; sc < colCombos.Count; sc++)
                        {
                            if (sc == c) continue;
                            if (dimIdx >= colCombos[sc].Count) continue;
                            bool sameContext = true;
                            for (int d = 0; d < colCombos[sc].Count; d++)
                            {
                                if (d == dimIdx) continue;
                                if (d >= colCombos[c].Count || colCombos[sc][d].Id != colCombos[c][d].Id)
                                { sameContext = false; break; }
                            }
                            if (sameContext)
                                siblingValues[colCombos[sc][dimIdx].Name] = result.Values[r, sc];
                        }
                    }

                    siblingValues[formulaMember.Name] = result.Values[r, c];

                    var normalizedFormula = ExpandSumRanges(NormalizeFormula(formulaMember.Formula!), formulaDimId);
                    var cellCtx = new Dictionary<long, long>(CurrentView!.PovSelections);
                    for (int d = 0; d < rowCombos[r].Count; d++)
                        cellCtx[CurrentView.RowAxes[d].DimensionId] = rowCombos[r][d].Id;
                    for (int d = 0; d < colCombos[c].Count; d++)
                        cellCtx[CurrentView.ColAxes[d].DimensionId] = colCombos[c][d].Id;
                    FillOffGridSiblings(formulaDimId, normalizedFormula, siblingValues, cellCtx, parsedFacts, descendantCache, dims);

                    result.Values[r, c] = EvaluateFormula(normalizedFormula, siblingValues);
                    continue;
                }

                // No on-axis formula member — check POV for a formula member.
                foreach (var (pid, pmid) in CurrentView!.PovSelections)
                {
                    var pm = _repo.GetMember(pmid);
                    if (pm != null && !string.IsNullOrEmpty(pm.Formula))
                    { formulaMember = pm; formulaDimId = pid; formulaMemberInPov = true; break; }
                }
                if (!formulaMemberInPov) continue;

                {
                    // All formula tokens are off-grid — FillOffGridSiblings aggregates each from facts.
                    var siblingValues = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
                    var normalizedFormula = ExpandSumRanges(NormalizeFormula(formulaMember!.Formula!), formulaDimId);
                    var cellCtx = new Dictionary<long, long>(CurrentView!.PovSelections);
                    for (int d = 0; d < rowCombos[r].Count; d++)
                        cellCtx[CurrentView.RowAxes[d].DimensionId] = rowCombos[r][d].Id;
                    for (int d = 0; d < colCombos[c].Count; d++)
                        cellCtx[CurrentView.ColAxes[d].DimensionId] = colCombos[c][d].Id;
                    FillOffGridSiblings(formulaDimId, normalizedFormula, siblingValues, cellCtx, parsedFacts, descendantCache, dims);
                    result.Values[r, c] = EvaluateFormula(normalizedFormula, siblingValues);
                }
            }
        }
    }

    public static decimal? EvaluateFormula(string formula, Dictionary<string, decimal?> memberValues)
    {
        var tokens = TokenizeFormula(formula);
        if (tokens.Count == 0) return null;

        var values = new List<decimal?>();
        var ops = new List<char>();

        for (int i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            if (token.Length == 1 && "+-*/".Contains(token[0]))
            {
                ops.Add(token[0]);
            }
            else if (token.StartsWith("\"") && token.EndsWith("\""))
            {
                var name = token[1..^1];
                if (memberValues.TryGetValue(name, out var v))
                    values.Add(v);
                else
                    return null;
            }
            else if (decimal.TryParse(token, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var num))
            {
                values.Add(num);
            }
            else
            {
                if (memberValues.TryGetValue(token, out var v))
                    values.Add(v);
                else
                    return null;
            }
        }

        if (values.Count == 0) return null;
        if (values.Any(v => !v.HasValue)) return null;

        var vals = values.Select(v => v!.Value).ToList();

        for (int i = 0; i < ops.Count; i++)
        {
            if (ops[i] == '*' || ops[i] == '/')
            {
                if (i + 1 >= vals.Count) break;
                decimal res = ops[i] == '*' ? vals[i] * vals[i + 1]
                    : (vals[i + 1] != 0 ? vals[i] / vals[i + 1] : 0);
                vals[i] = res;
                vals.RemoveAt(i + 1);
                ops.RemoveAt(i);
                i--;
            }
        }

        decimal result2 = vals[0];
        for (int i = 0; i < ops.Count; i++)
        {
            if (i + 1 >= vals.Count) break;
            result2 = ops[i] == '-' ? result2 - vals[i + 1] : result2 + vals[i + 1];
        }

        return result2;
    }

    // Returns leaf time members (ConsolOp="+", no children) in DFS/sort order for Sum(X:Y) expansion.
    private List<Member> GetLeavesInOrder(long dimId)
    {
        var all = _repo.GetMembers(dimId);
        var leaves = new List<Member>();
        void DFS(long? parentId)
        {
            foreach (var child in all.Where(m => m.ParentId == parentId).OrderBy(m => m.SortOrder))
            {
                if (child.ConsolOperator != "+" && child.ConsolOperator != "-") continue;
                if (all.Any(m => m.ParentId == child.Id))
                    DFS(child.Id);
                else
                    leaves.Add(child);
            }
        }
        DFS(null);
        return leaves;
    }

    // Expands Sum(X:Y) range notation to "X" + "m1" + ... + "Y" using leaf member order.
    private string ExpandSumRanges(string formula, long dimId)
    {
        if (!formula.Contains("Sum(", StringComparison.OrdinalIgnoreCase)) return formula;
        var leaves = GetLeavesInOrder(dimId);
        return Regex.Replace(formula, @"Sum\(([^:]+):([^)]+)\)", m =>
        {
            var startName = m.Groups[1].Value.Trim();
            var endName   = m.Groups[2].Value.Trim();
            int si = leaves.FindIndex(lf => lf.Name.Equals(startName, StringComparison.OrdinalIgnoreCase));
            int ei = leaves.FindIndex(lf => lf.Name.Equals(endName,   StringComparison.OrdinalIgnoreCase));
            if (si < 0 || ei < 0 || si > ei) return m.Value;
            return string.Join(" + ", leaves[si..(ei + 1)].Select(lf => $"\"{lf.Name}\""));
        }, RegexOptions.IgnoreCase);
    }

    private static List<string> TokenizeFormula(string formula)
    {
        var tokens = new List<string>();
        int i = 0;
        while (i < formula.Length)
        {
            if (char.IsWhiteSpace(formula[i])) { i++; continue; }
            if (formula[i] == '"')
            {
                int end = formula.IndexOf('"', i + 1);
                if (end < 0) end = formula.Length - 1;
                tokens.Add(formula[i..(end + 1)]);
                i = end + 1;
            }
            else if ("+-*/".Contains(formula[i]))
            {
                tokens.Add(formula[i].ToString());
                i++;
            }
            else
            {
                int start = i;
                while (i < formula.Length && !char.IsWhiteSpace(formula[i]) && !"+-*/\"".Contains(formula[i]))
                    i++;
                tokens.Add(formula[start..i]);
            }
        }
        return tokens;
    }

    // Strips label prefixes like "Profit Margin = expr" → "expr", and leading "=" → expr.
    private static string NormalizeFormula(string formula)
    {
        formula = formula.Trim();
        if (formula.StartsWith("=")) return formula[1..].Trim();
        // "Label" = expr  (spaces around =)
        int eq = formula.IndexOf(" = ", StringComparison.Ordinal);
        if (eq >= 0) return formula[(eq + 3)..].Trim();
        // "Label"=expr  (no spaces — e.g. "DailySales"="Sales"/"SalesDays")
        var m = System.Text.RegularExpressions.Regex.Match(formula, @"^""[^""]+""\s*=\s*(.+)$");
        if (m.Success) return m.Groups[1].Value.Trim();
        return formula;
    }

    // For formula tokens not visible on the grid, aggregate their value directly from facts.
    private void FillOffGridSiblings(long dimId, string formula, Dictionary<string, decimal?> siblingValues,
        Dictionary<long, long> cellCtx,
        List<(FactData fact, Dictionary<long, long> members)> parsedFacts,
        Dictionary<long, Dictionary<long, int>> descendantCache,
        List<Dimension> dims)
    {
        // Detect time-context type — only applies when this formula is NOT in the Time dimension itself.
        bool atOpBalTime = false;          // true OpBal (x-leaf, no formula) → flow/empty siblings = 0
        bool atFormulatedTimePeriod = false; // formula-based x-leaf (e.g. YTDFeb) → expand time formula for off-grid siblings
        Member? ctxTimeMember = null;
        var timeDim = dims.FirstOrDefault(d => d.DimType == DimensionType.Time);
        if (timeDim != null && dimId != timeDim.Id && cellCtx.TryGetValue(timeDim.Id, out var ctxTimeMemberId))
        {
            ctxTimeMember = _repo.GetMember(ctxTimeMemberId);
            // True OpBal: ConsolOperator=x, no formula, no children.
            // Formulated time member: has a formula regardless of ConsolOperator
            //   (e.g. YTDSep created with "+" still needs two-pass expansion).
            bool isLeafX = ctxTimeMember?.ConsolOperator == "x" &&
                string.IsNullOrEmpty(ctxTimeMember?.Formula) &&
                _repo.GetChildren(ctxTimeMemberId).Count == 0;
            if (isLeafX)
                atOpBalTime = true;
            else if (ctxTimeMember != null && !string.IsNullOrEmpty(ctxTimeMember.Formula))
                atFormulatedTimePeriod = true;
        }

        var tokens = TokenizeFormula(formula);
        var dimMembers = _repo.GetMembers(dimId);
        foreach (var token in tokens)
        {
            if (token.Length == 1 && "+-*/".Contains(token[0])) continue;
            var name = token.StartsWith("\"") && token.EndsWith("\"") ? token[1..^1] : token;
            if (siblingValues.ContainsKey(name)) continue;
            if (decimal.TryParse(name, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out _)) continue;

            var refMember = dimMembers.FirstOrDefault(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (refMember == null) continue;

            if (atOpBalTime)
            {
                // True OpBal: flow/empty-TimeBalance off-grid members have 0 as opening balance
                var refTb = refMember.TimeBalance?.Trim().ToLowerInvariant() ?? "";
                if (refTb == "flow" || refTb == "")
                { siblingValues[name] = 0m; continue; }
            }
            else if (atFormulatedTimePeriod)
            {
                // Two-pass YTD: expand the time formula for the off-grid measure sibling.
                // e.g., Sales at YTDAug: apply YTDAug formula "YTDJul"+"Aug" recursively.
                // This allows Margin at YTDAug = (TP_Jan+...+TP_Aug) / (Sales_Jan+...+Sales_Aug) * 100.
                var timeFormula = ExpandSumRanges(NormalizeFormula(ctxTimeMember!.Formula!), timeDim!.Id);
                var timeTokens = TokenizeFormula(timeFormula);
                var allTimeMembers = _repo.GetMembers(timeDim!.Id);
                var timeSibValues = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
                var baseCtxTP = new Dictionary<long, long>(cellCtx) { [dimId] = refMember.Id };
                foreach (var tt in timeTokens)
                {
                    if (tt.Length == 1 && "+-*/".Contains(tt[0])) continue;
                    var tname = tt.StartsWith("\"") && tt.EndsWith("\"") ? tt[1..^1] : tt;
                    if (decimal.TryParse(tname, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out _)) continue;
                    if (timeSibValues.ContainsKey(tname)) continue;
                    // Try originals first, then shared copies, to handle facts loaded under any copy.
                    decimal? tval = null;
                    foreach (var cand in allTimeMembers
                        .Where(m => m.Name.Equals(tname, StringComparison.OrdinalIgnoreCase))
                        .OrderBy(m => m.SharedFromId.HasValue ? 1 : 0).ThenBy(m => m.Id))
                    {
                        tval = EvaluateTimeMemberValue(cand.Id, baseCtxTP, parsedFacts, descendantCache, timeDim.Id, allTimeMembers);
                        if (tval.HasValue) break;
                    }
                    timeSibValues[tname] = tval;
                }
                siblingValues[name] = EvaluateFormula(timeFormula, timeSibValues);
                continue;
            }

            // When evaluating a time dimension formula (dimId == timeDim.Id), use recursive
            // time evaluation so formula chains like YTDAug→YTDJul→YTDJun→months resolve correctly.
            if (timeDim != null && dimId == timeDim.Id)
            {
                decimal? tval = null;
                foreach (var cand in dimMembers
                    .Where(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(m => m.SharedFromId.HasValue ? 1 : 0).ThenBy(m => m.Id))
                {
                    tval = EvaluateTimeMemberValue(cand.Id, cellCtx, parsedFacts, descendantCache, dimId, dimMembers);
                    if (tval.HasValue) break;
                }
                siblingValues[name] = tval;
                continue;
            }

            EnsureDescendantsForMember(refMember.Id, descendantCache);
            var ctx = new Dictionary<long, long>(cellCtx) { [dimId] = refMember.Id };
            siblingValues[name] = AggregateFromFacts(ctx, parsedFacts, descendantCache);
        }
    }

    // Recursively evaluates a time member's aggregated value for a given context, following formula chains.
    // baseCtx has all non-time dimension slots filled; the time slot is set per recursive call.
    // Tries original members (SharedFromId==null) before shared copies so facts loaded under any
    // copy of a month are found without double-counting.
    private decimal? EvaluateTimeMemberValue(
        long timeMemberId,
        Dictionary<long, long> baseCtx,
        List<(FactData fact, Dictionary<long, long> members)> parsedFacts,
        Dictionary<long, Dictionary<long, int>> descendantCache,
        long timeDimId,
        List<Member> allTimeMembers,
        int depth = 0)
    {
        const int MaxDepth = 25;
        if (depth >= MaxDepth) return null;

        var member = allTimeMembers.FirstOrDefault(m => m.Id == timeMemberId);

        // Formula members always evaluate via formula (overrides child rollup).
        if (member != null && !string.IsNullOrEmpty(member.Formula))
        {
            var normalizedFormula = ExpandSumRanges(NormalizeFormula(member.Formula), timeDimId);
            var tokens = TokenizeFormula(normalizedFormula);
            var formulaValues = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
            foreach (var token in tokens)
            {
                if (token.Length == 1 && "+-*/".Contains(token[0])) continue;
                var name = token.StartsWith("\"") && token.EndsWith("\"") ? token[1..^1] : token;
                if (decimal.TryParse(name, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out _)) continue;
                if (formulaValues.ContainsKey(name)) continue;
                // Try originals first, fall back to shared copies.
                decimal? tokenVal = null;
                foreach (var cand in allTimeMembers
                    .Where(m => m.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(m => m.SharedFromId.HasValue ? 1 : 0).ThenBy(m => m.Id))
                {
                    tokenVal = EvaluateTimeMemberValue(cand.Id, baseCtx, parsedFacts, descendantCache, timeDimId, allTimeMembers, depth + 1);
                    if (tokenVal.HasValue) break;
                }
                formulaValues[name] = tokenVal;
            }
            return EvaluateFormula(normalizedFormula, formulaValues);
        }

        // Non-formula: aggregate directly from facts (handles leaf months and non-formula rollups).
        EnsureDescendantsForMember(timeMemberId, descendantCache);
        var ctx = new Dictionary<long, long>(baseCtx) { [timeDimId] = timeMemberId };
        return AggregateFromFacts(ctx, parsedFacts, descendantCache);
    }

    private void EnsureDescendantsForMember(long memberId, Dictionary<long, Dictionary<long, int>> cache)
    {
        if (cache.ContainsKey(memberId)) return;
        var signs = new Dictionary<long, int> { [memberId] = 1 };
        void Build(long pid, int ps)
        {
            foreach (var ch in _repo.GetChildren(pid))
            {
                int cs = ch.ConsolOperator == "x" ? 0 : ch.ConsolOperator == "-" ? ps * -1 : ps;
                var eff = ch.SharedFromId ?? ch.Id;
                if (!signs.ContainsKey(eff)) signs[eff] = cs;
                signs[ch.Id] = cs;
                Build(ch.Id, cs);
            }
        }
        // If this member is a shared copy, include the original and its full sub-hierarchy.
        var rootMem = _repo.GetMember(memberId);
        if (rootMem?.SharedFromId != null)
        {
            signs.TryAdd(rootMem.SharedFromId.Value, 1);
            Build(rootMem.SharedFromId.Value, 1);
        }
        Build(memberId, 1);
        cache[memberId] = signs;
    }

    private static decimal? AggregateFromFacts(Dictionary<long, long> cellCtx,
        List<(FactData fact, Dictionary<long, long> members)> parsedFacts,
        Dictionary<long, Dictionary<long, int>> descendantCache)
    {
        decimal total = 0; bool any = false;
        foreach (var (fact, factMembers) in parsedFacts)
        {
            if (!fact.NumericValue.HasValue) continue;
            bool match = true; int netSign = 1;
            foreach (var (dimId, targetId) in cellCtx)
            {
                if (!factMembers.TryGetValue(dimId, out var fmId)) continue;
                if (!descendantCache.TryGetValue(targetId, out var smap) ||
                    !smap.TryGetValue(fmId, out var sign)) { match = false; break; }
                if (sign == 0) { match = false; break; }
                netSign *= sign;
            }
            if (match) { total += fact.NumericValue.Value * netSign; any = true; }
        }
        return any ? total : null;
    }

    private Member GetLastLeafDescendant(Member member)
    {
        var children = _repo.GetChildren(member.Id);
        if (children.Count == 0) return member;
        return GetLastLeafDescendant(children.Last());
    }

    private void ApplyTimeBalance(GridResult result, List<List<Member>> rowCombos, List<List<Member>> colCombos, List<Dimension> dims,
        List<(FactData fact, Dictionary<long, long> members)> parsedFacts, Dictionary<long, Dictionary<long, int>> descendantCache)
    {
        if (CurrentView == null) return;

        int timeDimRowIdx = -1;
        int timeDimColIdx = -1;
        int measureDimRowIdx = -1;
        int measureDimColIdx = -1;

        for (int i = 0; i < CurrentView.RowAxes.Count; i++)
        {
            var dim = dims.FirstOrDefault(d => d.Id == CurrentView.RowAxes[i].DimensionId);
            if (dim?.DimType == DimensionType.Time) timeDimRowIdx = i;
            if (dim?.DimType == DimensionType.Measure) measureDimRowIdx = i;
        }
        for (int i = 0; i < CurrentView.ColAxes.Count; i++)
        {
            var dim = dims.FirstOrDefault(d => d.Id == CurrentView.ColAxes[i].DimensionId);
            if (dim?.DimType == DimensionType.Time) timeDimColIdx = i;
            if (dim?.DimType == DimensionType.Measure) measureDimColIdx = i;
        }

        if (timeDimRowIdx < 0 && timeDimColIdx < 0) return;

        for (int r = 0; r < rowCombos.Count; r++)
        {
            for (int c = 0; c < colCombos.Count; c++)
            {
                Member? timeMember = null;
                bool timeOnRow = false;
                if (timeDimRowIdx >= 0 && timeDimRowIdx < rowCombos[r].Count)
                { timeMember = rowCombos[r][timeDimRowIdx]; timeOnRow = true; }
                else if (timeDimColIdx >= 0 && timeDimColIdx < colCombos[c].Count)
                    timeMember = colCombos[c][timeDimColIdx];

                if (timeMember == null) continue;
                var timeChildren = _repo.GetChildren(timeMember.Id);

                // OpBal and similar: leaf Time member excluded from rollup (ConsolOp=x, no children, no formula, no fact data).
                // Formula-based x-leaf members (e.g. YTDFeb) are NOT OpBal — their values come from ApplyFormulas.
                // For true OpBal, flow/empty-TimeBalance measures have 0 as opening balance.
                if (timeMember.ConsolOperator == "x" && timeChildren.Count == 0 && string.IsNullOrEmpty(timeMember.Formula))
                {
                    Member? opMeasure = null;
                    if (measureDimRowIdx >= 0 && measureDimRowIdx < rowCombos[r].Count)
                        opMeasure = rowCombos[r][measureDimRowIdx];
                    else if (measureDimColIdx >= 0 && measureDimColIdx < colCombos[c].Count)
                        opMeasure = colCombos[c][measureDimColIdx];
                    if (opMeasure != null)
                    {
                        var opTb = opMeasure.TimeBalance?.Trim().ToLowerInvariant() ?? "";
                        if (opTb == "flow" || opTb == "")
                            result.Values[r, c] = 0m;
                    }
                    continue;
                }

                if (timeChildren.Count == 0) continue;

                Member? measureMember = null;
                if (measureDimRowIdx >= 0 && measureDimRowIdx < rowCombos[r].Count)
                    measureMember = rowCombos[r][measureDimRowIdx];
                else if (measureDimColIdx >= 0 && measureDimColIdx < colCombos[c].Count)
                    measureMember = colCombos[c][measureDimColIdx];

                if (measureMember == null) continue;
                if (string.IsNullOrEmpty(measureMember.TimeBalance)) continue;

                var tb = measureMember.TimeBalance.Trim().ToLowerInvariant();
                if (tb == "flow") continue;

                if (tb == "last")
                {
                    var lastChild = timeChildren.Last();
                    decimal? lastVal = null;
                    bool foundInGrid = false;
                    if (timeOnRow)
                    {
                        for (int sr = 0; sr < rowCombos.Count; sr++)
                        {
                            if (timeDimRowIdx >= rowCombos[sr].Count) continue;
                            if (rowCombos[sr][timeDimRowIdx].Id == lastChild.Id)
                            {
                                bool sameContext = true;
                                for (int d = 0; d < rowCombos[sr].Count; d++)
                                {
                                    if (d == timeDimRowIdx) continue;
                                    if (d >= rowCombos[r].Count || rowCombos[sr][d].Id != rowCombos[r][d].Id)
                                    { sameContext = false; break; }
                                }
                                if (sameContext) { lastVal = result.Values[sr, c]; foundInGrid = true; break; }
                            }
                        }
                    }
                    else
                    {
                        for (int sc = 0; sc < colCombos.Count; sc++)
                        {
                            if (timeDimColIdx >= colCombos[sc].Count) continue;
                            if (colCombos[sc][timeDimColIdx].Id == lastChild.Id)
                            {
                                bool sameContext = true;
                                for (int d = 0; d < colCombos[sc].Count; d++)
                                {
                                    if (d == timeDimColIdx) continue;
                                    if (d >= colCombos[c].Count || colCombos[sc][d].Id != colCombos[c][d].Id)
                                    { sameContext = false; break; }
                                }
                                if (sameContext) { lastVal = result.Values[r, sc]; foundInGrid = true; break; }
                            }
                        }
                    }
                    if (!foundInGrid)
                    {
                        // Last child not visible — walk to the deepest last leaf and look up directly
                        var lastLeaf = GetLastLeafDescendant(lastChild);
                        if (!descendantCache.ContainsKey(lastLeaf.Id))
                        {
                            var leafMap = new Dictionary<long, int> { [lastLeaf.Id] = 1 };
                            // Shared member: facts are keyed by the original member's ID, not the copy's.
                            if (lastLeaf.SharedFromId.HasValue)
                                leafMap[lastLeaf.SharedFromId.Value] = 1;
                            descendantCache[lastLeaf.Id] = leafMap;
                        }
                        var cellMembers = new Dictionary<long, long>(CurrentView!.PovSelections);
                        for (int d = 0; d < rowCombos[r].Count; d++)
                        {
                            var axisId = CurrentView.RowAxes[d].DimensionId;
                            cellMembers[axisId] = (d == timeDimRowIdx) ? lastLeaf.Id : rowCombos[r][d].Id;
                        }
                        for (int d = 0; d < colCombos[c].Count; d++)
                        {
                            var axisId = CurrentView.ColAxes[d].DimensionId;
                            cellMembers[axisId] = (d == timeDimColIdx) ? lastLeaf.Id : colCombos[c][d].Id;
                        }
                        decimal leafTotal = 0; bool leafHasValue = false;
                        foreach (var (fact, factMembers) in parsedFacts)
                        {
                            if (!fact.NumericValue.HasValue) continue;
                            bool match = true; int netSign = 1;
                            foreach (var (dimId, targetId) in cellMembers)
                            {
                                if (!factMembers.TryGetValue(dimId, out var factMemberId)) continue;
                                if (!descendantCache.TryGetValue(targetId, out var signMap) ||
                                    !signMap.TryGetValue(factMemberId, out var sign)) { match = false; break; }
                                if (sign == 0) { match = false; break; }
                                netSign *= sign;
                            }
                            if (match) { leafTotal += fact.NumericValue.Value * netSign; leafHasValue = true; }
                        }
                        lastVal = leafHasValue ? leafTotal : null;
                    }
                    result.Values[r, c] = lastVal;
                }
                else if (tb == "equation" && !string.IsNullOrEmpty(measureMember.Formula))
                {
                    var siblingValues = new Dictionary<string, decimal?>(StringComparer.OrdinalIgnoreCase);
                    int mIdx = measureDimRowIdx >= 0 ? measureDimRowIdx : -1;
                    bool mOnRow = measureDimRowIdx >= 0;

                    if (mOnRow)
                    {
                        for (int sr = 0; sr < rowCombos.Count; sr++)
                        {
                            if (mIdx >= rowCombos[sr].Count) continue;
                            bool sameContext = true;
                            for (int d = 0; d < rowCombos[sr].Count; d++)
                            {
                                if (d == mIdx) continue;
                                if (d >= rowCombos[r].Count || rowCombos[sr][d].Id != rowCombos[r][d].Id)
                                { sameContext = false; break; }
                            }
                            if (sameContext)
                                siblingValues[rowCombos[sr][mIdx].Name] = result.Values[sr, c];
                        }
                    }
                    else
                    {
                        int mcIdx = measureDimColIdx;
                        for (int sc = 0; sc < colCombos.Count; sc++)
                        {
                            if (mcIdx >= colCombos[sc].Count) continue;
                            bool sameContext = true;
                            for (int d = 0; d < colCombos[sc].Count; d++)
                            {
                                if (d == mcIdx) continue;
                                if (d >= colCombos[c].Count || colCombos[sc][d].Id != colCombos[c][d].Id)
                                { sameContext = false; break; }
                            }
                            if (sameContext)
                                siblingValues[colCombos[sc][mcIdx].Name] = result.Values[r, sc];
                        }
                    }
                    bool mOnRowForCtx = measureDimRowIdx >= 0;
                    int mIdxForCtx = mOnRowForCtx ? measureDimRowIdx : measureDimColIdx;
                    long measureDimIdForCtx = mOnRowForCtx
                        ? CurrentView!.RowAxes[mIdxForCtx].DimensionId
                        : CurrentView!.ColAxes[mIdxForCtx].DimensionId;
                    var cellCtxTb = new Dictionary<long, long>(CurrentView!.PovSelections);
                    for (int d = 0; d < rowCombos[r].Count; d++)
                        cellCtxTb[CurrentView.RowAxes[d].DimensionId] = rowCombos[r][d].Id;
                    for (int d = 0; d < colCombos[c].Count; d++)
                        cellCtxTb[CurrentView.ColAxes[d].DimensionId] = colCombos[c][d].Id;
                    var normalizedMFormula = ExpandSumRanges(NormalizeFormula(measureMember.Formula), measureDimIdForCtx);
                    FillOffGridSiblings(measureDimIdForCtx, normalizedMFormula, siblingValues, cellCtxTb, parsedFacts, descendantCache, dims);
                    result.Values[r, c] = EvaluateFormula(normalizedMFormula, siblingValues);
                }
            }
        }
    }

    /// <summary>
    /// Builds the pipe-delimited member key used to look up fact values.
    /// Key order matches dimension SortOrder.
    /// </summary>
    public static string BuildMemberKey(List<long> dimOrder, Dictionary<long, long> memberIds)
    {
        var parts = new List<string>();
        foreach (var dimId in dimOrder)
        {
            if (memberIds.TryGetValue(dimId, out var memberId))
                parts.Add(memberId.ToString());
            else
                parts.Add("0");
        }
        return string.Join("|", parts);
    }

    #region OLAP Operations

    /// <summary>
    /// Pushes current view to undo stack before making changes.
    /// </summary>
    private void PushUndo()
    {
        if (CurrentView != null)
            _undo.Push(CurrentView);
    }

    public ViewState? Undo()
    {
        var prev = _undo.Pop();
        if (prev != null)
            CurrentView = prev;
        return CurrentView;
    }

    public bool CanUndo => _undo.CanUndo;
    public bool UndoLimitReached => _undo.LimitReached;
    public int UndoTotalPushed => _undo.TotalPushed;

    /// <summary>
    /// Drill down on a member: replace it with its children in the view.
    /// </summary>
    public void DrillDown(long dimensionId, long memberId, DrillMode mode)
    {
        if (CurrentView == null) return;
        PushUndo();

        List<long> replacementIds;
        switch (mode)
        {
            case DrillMode.NextGeneration:
                replacementIds = _repo.GetChildren(memberId).Select(m => m.Id).ToList();
                break;
            case DrillMode.AllGenerations:
                replacementIds = GetDescendantsPostOrder(memberId);
                break;
            case DrillMode.BaseOnly:
                replacementIds = _repo.GetLeafDescendants(memberId).Select(m => m.Id).ToList();
                break;
            default:
                return;
        }

        if (replacementIds.Count == 0) return;
        // Rollup (parent) appears after its children so totals sit below detail rows.
        if (mode == DrillMode.AllGenerations || mode == DrillMode.NextGeneration) replacementIds.Add(memberId);

        // If children are already visible (e.g. from a previous drill with the old parent-first order),
        // pre-remove them so ReplaceInAxes deduplication does not silently drop them from the
        // replacement list, which would leave the parent in its original position.
        var childrenToPreRemove = replacementIds.Where(id => id != memberId).ToList();
        foreach (var childId in childrenToPreRemove)
        {
            foreach (var axis in CurrentView.RowAxes.Where(a => a.DimensionId == dimensionId))
                axis.VisibleMemberIds.Remove(childId);
            foreach (var axis in CurrentView.ColAxes.Where(a => a.DimensionId == dimensionId))
                axis.VisibleMemberIds.Remove(childId);
        }

        ReplaceInAxes(CurrentView.RowAxes, dimensionId, memberId, replacementIds);
        ReplaceInAxes(CurrentView.ColAxes, dimensionId, memberId, replacementIds);
    }

    /// <summary>
    /// Drill up: replace the member with its parent (or go to root).
    /// </summary>
    public void DrillUp(long dimensionId, long memberId)
    {
        if (CurrentView == null) return;
        PushUndo();

        var member = _repo.GetMember(memberId);
        if (member?.ParentId == null)
        {
            var roots = _repo.GetRootMembers(dimensionId);
            var rootIds = roots.Select(r => r.Id).ToList();
            ReplaceAllInAxis(CurrentView.RowAxes, dimensionId, rootIds);
            ReplaceAllInAxis(CurrentView.ColAxes, dimensionId, rootIds);
            return;
        }

        var siblings = FindSiblingsInView(dimensionId, memberId);
        foreach (var sibId in siblings)
        {
            ReplaceInAxes(CurrentView.RowAxes, dimensionId, sibId, new List<long> { member.ParentId.Value });
            ReplaceInAxes(CurrentView.ColAxes, dimensionId, sibId, new List<long> { member.ParentId.Value });
        }
    }

    /// <summary>
    /// Swap row and column axes (pivot).
    /// </summary>
    public void SwapRowCol()
    {
        if (CurrentView == null) return;
        PushUndo();
        (CurrentView.RowAxes, CurrentView.ColAxes) = (CurrentView.ColAxes, CurrentView.RowAxes);
    }
    public void SwapDimension(long dimensionId)
    {
        if (CurrentView == null) return;
        PushUndo();

        var rowAxis = CurrentView.RowAxes.FirstOrDefault(a => a.DimensionId == dimensionId);
        var colAxis = CurrentView.ColAxes.FirstOrDefault(a => a.DimensionId == dimensionId);

        if (rowAxis != null)
        {
            CurrentView.RowAxes.Remove(rowAxis);
            EnsureAxesNonEmpty();
            CurrentView.ColAxes.Add(rowAxis);
        }
        else if (colAxis != null)
        {
            CurrentView.ColAxes.Remove(colAxis);
            EnsureAxesNonEmpty();
            CurrentView.RowAxes.Add(colAxis);
        }
        else if (CurrentView.PovSelections.TryGetValue(dimensionId, out var povMemberId))
        {
            CurrentView.PovSelections.Remove(dimensionId);
            var dim = _repo.GetDimensions(CurrentView.ModelId).FirstOrDefault(d => d.Id == dimensionId);
            var roots = _repo.GetRootMembers(dimensionId);
            CurrentView.ColAxes.Add(new DimensionAxis
            {
                DimensionId = dimensionId,
                DimensionName = dim?.Name ?? "Unknown",
                VisibleMemberIds = roots.Select(r => r.Id).ToList()
            });
        }
    }

    /// <summary>
    /// Move the dimension to the row axis (from col axis or POV).
    /// Returns false if the dimension is already on the row axis.
    /// </summary>
    public bool MoveToRow(long dimensionId)
    {
        if (CurrentView == null) return false;
        if (CurrentView.RowAxes.Any(a => a.DimensionId == dimensionId)) return false;
        PushUndo();

        var colAxis = CurrentView.ColAxes.FirstOrDefault(a => a.DimensionId == dimensionId);
        if (colAxis != null)
        {
            CurrentView.ColAxes.Remove(colAxis);
            EnsureAxesNonEmpty();
            CurrentView.RowAxes.Add(colAxis);
            return true;
        }
        if (CurrentView.PovSelections.TryGetValue(dimensionId, out var povMemberId))
        {
            CurrentView.PovSelections.Remove(dimensionId);
            var dim = _repo.GetDimensions(CurrentView.ModelId).FirstOrDefault(d => d.Id == dimensionId);
            CurrentView.RowAxes.Add(new DimensionAxis
            {
                DimensionId = dimensionId,
                DimensionName = dim?.Name ?? "Unknown",
                VisibleMemberIds = new List<long> { povMemberId }
            });
            return true;
        }
        return false;
    }

    /// <summary>
    /// Move the dimension to the column axis (from row axis or POV).
    /// Returns false if the dimension is already on the column axis.
    /// </summary>
    public bool MoveToCol(long dimensionId)
    {
        if (CurrentView == null) return false;
        if (CurrentView.ColAxes.Any(a => a.DimensionId == dimensionId)) return false;
        PushUndo();

        var rowAxis = CurrentView.RowAxes.FirstOrDefault(a => a.DimensionId == dimensionId);
        if (rowAxis != null)
        {
            CurrentView.RowAxes.Remove(rowAxis);
            EnsureAxesNonEmpty();
            CurrentView.ColAxes.Add(rowAxis);
            return true;
        }
        if (CurrentView.PovSelections.TryGetValue(dimensionId, out var povMemberId))
        {
            CurrentView.PovSelections.Remove(dimensionId);
            var dim = _repo.GetDimensions(CurrentView.ModelId).FirstOrDefault(d => d.Id == dimensionId);
            CurrentView.ColAxes.Add(new DimensionAxis
            {
                DimensionId = dimensionId,
                DimensionName = dim?.Name ?? "Unknown",
                VisibleMemberIds = new List<long> { povMemberId }
            });
            return true;
        }
        return false;
    }

    /// <summary>
    /// Keep only the selected member, remove all others in the same dimension.
    /// </summary>
    public void KeepSelected(long dimensionId, long memberId)
    {
        if (CurrentView == null) return;
        PushUndo();

        foreach (var axis in CurrentView.RowAxes.Where(a => a.DimensionId == dimensionId))
            axis.VisibleMemberIds = new List<long> { memberId };
        foreach (var axis in CurrentView.ColAxes.Where(a => a.DimensionId == dimensionId))
            axis.VisibleMemberIds = new List<long> { memberId };
    }

    /// <summary>
    /// Remove only the selected member from the view.
    /// If removing the last visible member on an axis, resets to root members so the axis stays populated.
    /// </summary>
    public void RemoveSelected(long dimensionId, long memberId)
    {
        if (CurrentView == null) return;
        PushUndo();

        foreach (var axis in CurrentView.RowAxes.Where(a => a.DimensionId == dimensionId))
        {
            axis.VisibleMemberIds.Remove(memberId);
            if (axis.VisibleMemberIds.Count == 0)
                axis.VisibleMemberIds = _repo.GetRootMembers(dimensionId).Select(m => m.Id).ToList();
        }
        foreach (var axis in CurrentView.ColAxes.Where(a => a.DimensionId == dimensionId))
        {
            axis.VisibleMemberIds.Remove(memberId);
            if (axis.VisibleMemberIds.Count == 0)
                axis.VisibleMemberIds = _repo.GetRootMembers(dimensionId).Select(m => m.Id).ToList();
        }
    }

    /// <summary>
    /// Places a picked member onto the row or column axis.
    /// </summary>
    public void PickMember(long dimensionId, long memberId, bool onRow)
    {
        PickMembers(dimensionId, new List<long> { memberId }, onRow);
    }

    /// <summary>
    /// Places multiple picked members onto the row or column axis,
    /// replacing any previous members for that dimension on the target axis.
    /// </summary>
    public void PickMembers(long dimensionId, List<long> memberIds, bool onRow)
    {
        if (CurrentView == null || memberIds.Count == 0) return;
        PushUndo();

        var targetAxes = onRow ? CurrentView.RowAxes : CurrentView.ColAxes;
        var otherAxes = onRow ? CurrentView.ColAxes : CurrentView.RowAxes;

        otherAxes.RemoveAll(a => a.DimensionId == dimensionId);
        CurrentView.PovSelections.Remove(dimensionId);
        EnsureAxesNonEmpty();

        var deduped = memberIds.Distinct().ToList();
        var existing = targetAxes.FirstOrDefault(a => a.DimensionId == dimensionId);
        if (existing != null)
        {
            existing.VisibleMemberIds = deduped;
        }
        else
        {
            var dim = _repo.GetDimensions(CurrentView.ModelId)
                .FirstOrDefault(d => d.Id == dimensionId);
            targetAxes.Add(new DimensionAxis
            {
                DimensionId = dimensionId,
                DimensionName = dim?.Name ?? "Unknown",
                VisibleMemberIds = deduped
            });
        }
    }

    /// <summary>
    /// Toggles a dimension between axis and POV header.
    /// Axis → POV: collapses to the hinted member (the cell the user clicked), or first visible member.
    /// POV → axis: expands all root members onto the row axis.
    /// </summary>
    /// <summary>
    /// Returns false if the move is blocked (too few grid dimensions remain). Caller should alert the user.
    /// Returns true if already in header (silent no-op) or if the move was executed successfully.
    /// </summary>
    public bool MoveToHeader(long dimensionId, long hintMemberId = 0)
    {
        if (CurrentView == null) return true;

        var axis = CurrentView.RowAxes.FirstOrDefault(a => a.DimensionId == dimensionId)
                ?? CurrentView.ColAxes.FirstOrDefault(a => a.DimensionId == dimensionId);

        // Already in POV header — silent no-op.
        if (axis == null) return true;

        // Need at least 2 dimensions remaining on axes so both rows and columns stay populated.
        int remainingAxesDims = CurrentView.RowAxes.Count + CurrentView.ColAxes.Count - 1;
        if (remainingAxesDims < 2) return false;

        PushUndo();
        // Axis → POV: use the hinted member (cell the user clicked), else first visible, else root.
        CurrentView.RowAxes.Remove(axis);
        CurrentView.ColAxes.Remove(axis);
        EnsureAxesNonEmpty();
        long povMemberId = 0;
        if (hintMemberId > 0 && axis.VisibleMemberIds.Contains(hintMemberId))
            povMemberId = hintMemberId;
        else if (axis.VisibleMemberIds.Count > 0)
            povMemberId = axis.VisibleMemberIds[0];
        else
            povMemberId = _repo.GetRootMembers(dimensionId).FirstOrDefault()?.Id ?? 0;
        if (povMemberId > 0)
            CurrentView.PovSelections[dimensionId] = povMemberId;
        return true;
    }

    #endregion

    // Ensures both axes always have at least one dimension. If one axis is empty and the other
    // has more than one, rescues the first dimension from the non-empty axis into the empty one.
    private void EnsureAxesNonEmpty()
    {
        if (CurrentView == null) return;
        if (CurrentView.RowAxes.Count == 0 && CurrentView.ColAxes.Count > 0)
        {
            var rescued = CurrentView.ColAxes[0];
            CurrentView.ColAxes.RemoveAt(0);
            CurrentView.RowAxes.Add(rescued);
        }
        else if (CurrentView.ColAxes.Count == 0 && CurrentView.RowAxes.Count > 0)
        {
            var rescued = CurrentView.RowAxes[0];
            CurrentView.RowAxes.RemoveAt(0);
            CurrentView.ColAxes.Add(rescued);
        }
    }

    #region Private helpers

    // DFS post-order: each node's children (recursively) come before the node itself,
    // so intermediate rollup rows appear after their detail rows in the grid.
    private List<long> GetDescendantsPostOrder(long memberId)
    {
        var result = new List<long>();
        foreach (var child in _repo.GetChildren(memberId))
        {
            result.AddRange(GetDescendantsPostOrder(child.Id));
            result.Add(child.Id);
        }
        return result;
    }

    private static void ReplaceInAxes(List<DimensionAxis> axes, long dimId, long oldId, List<long> newIds)
    {
        foreach (var axis in axes.Where(a => a.DimensionId == dimId))
        {
            var idx = axis.VisibleMemberIds.IndexOf(oldId);
            if (idx < 0) continue;
            axis.VisibleMemberIds.RemoveAt(idx);
            var deduplicated = newIds.Where(id => !axis.VisibleMemberIds.Contains(id)).ToList();
            axis.VisibleMemberIds.InsertRange(idx, deduplicated);
        }
    }

    private static void ReplaceAllInAxis(List<DimensionAxis> axes, long dimId, List<long> newIds)
    {
        foreach (var axis in axes.Where(a => a.DimensionId == dimId))
            axis.VisibleMemberIds = new List<long>(newIds.Distinct());
    }

    private List<long> FindSiblingsInView(long dimensionId, long memberId)
    {
        if (CurrentView == null) return new List<long> { memberId };
        var member = _repo.GetMember(memberId);
        if (member?.ParentId == null) return new List<long> { memberId };

        var allAxes = CurrentView.RowAxes.Concat(CurrentView.ColAxes);
        var axis = allAxes.FirstOrDefault(a => a.DimensionId == dimensionId);
        if (axis == null) return new List<long> { memberId };

        var siblings = _repo.GetChildren(member.ParentId.Value);
        var siblingIds = new HashSet<long>(siblings.Select(s => s.Id));
        return axis.VisibleMemberIds.Where(id => siblingIds.Contains(id)).ToList();
    }

    #endregion

    /// <summary>
    /// Computes the cartesian product of member lists (one per dimension on an axis).
    /// </summary>
    /// <summary>
    /// Reorders members so children appear before their parents (bottom-up / post-order).
    /// Parents act as subtotals at the bottom of their group.
    /// Falls back to original order if anything goes wrong.
    /// </summary>
    private static List<Member> ReorderBottomUp(List<Member> members)
    {
        if (members.Count <= 1) return members;

        try
        {
            var idSet = new HashSet<long>(members.Select(m => m.Id));
            var childrenOf = new Dictionary<long, List<Member>>();
            var roots = new List<Member>();

            foreach (var m in members)
            {
                if (m.ParentId.HasValue && idSet.Contains(m.ParentId.Value))
                {
                    if (!childrenOf.ContainsKey(m.ParentId.Value))
                        childrenOf[m.ParentId.Value] = new List<Member>();
                    childrenOf[m.ParentId.Value].Add(m);
                }
                else
                {
                    roots.Add(m);
                }
            }

            if (childrenOf.Count == 0) return members;

            var result = new List<Member>();
            var visited = new HashSet<long>();

            void PostOrder(Member m)
            {
                if (!visited.Add(m.Id)) return;
                if (childrenOf.TryGetValue(m.Id, out var kids))
                    foreach (var kid in kids)
                        PostOrder(kid);
                result.Add(m);
            }

            foreach (var root in roots)
                PostOrder(root);

            return result.Count == members.Count ? result : members;
        }
        catch
        {
            return members;
        }
    }

    private static List<List<Member>> CartesianProduct(List<List<Member>> lists)
    {
        if (lists.Count == 0)
            return new List<List<Member>>();

        if (lists.Any(l => l.Count == 0))
            return new List<List<Member>>();

        var result = new List<List<Member>> { new() };
        foreach (var list in lists)
        {
            var temp = new List<List<Member>>();
            foreach (var existing in result)
            {
                foreach (var item in list)
                {
                    var copy = new List<Member>(existing) { item };
                    temp.Add(copy);
                }
            }
            result = temp;
        }
        return result;
    }
}

public enum DrillMode
{
    NextGeneration,
    AllGenerations,
    BaseOnly
}

/// <summary>
/// The 2D grid produced by BuildGrid, ready to be written into the worksheet.
/// </summary>
public class GridResult
{
    public List<string> RowDimensionNames { get; set; } = new();
    public List<string> ColDimensionNames { get; set; } = new();
    public List<List<Member>> RowHeaders { get; set; } = new();
    public List<List<Member>> ColHeaders { get; set; } = new();
    public decimal?[,] Values { get; set; } = new decimal?[0, 0];
    public int MemberDisplay { get; set; }

    public string FormatMember(Member m)
    {
        return MemberDisplay switch
        {
            1 => m.Description,
            2 => $"{m.DisplayName} - {m.Description}",
            _ => m.DisplayName
        };
    }

    public void OmitEmptyRows()
    {
        var keep = new List<int>();
        for (int r = 0; r < RowHeaders.Count; r++)
        {
            bool hasValue = false;
            for (int c = 0; c < ColHeaders.Count; c++)
            {
                if (Values[r, c].HasValue) { hasValue = true; break; }
            }
            if (hasValue) keep.Add(r);
        }
        Reindex(keep, isRow: true);
    }

    public void OmitEmptyColumns()
    {
        var keep = new List<int>();
        for (int c = 0; c < ColHeaders.Count; c++)
        {
            bool hasValue = false;
            for (int r = 0; r < RowHeaders.Count; r++)
            {
                if (Values[r, c].HasValue) { hasValue = true; break; }
            }
            if (hasValue) keep.Add(c);
        }
        Reindex(keep, isRow: false);
    }

    private void Reindex(List<int> keep, bool isRow)
    {
        if (isRow)
        {
            var newRows = keep.Select(i => RowHeaders[i]).ToList();
            var newVals = new decimal?[keep.Count, ColHeaders.Count];
            for (int r = 0; r < keep.Count; r++)
                for (int c = 0; c < ColHeaders.Count; c++)
                    newVals[r, c] = Values[keep[r], c];
            RowHeaders = newRows;
            Values = newVals;
        }
        else
        {
            var newCols = keep.Select(i => ColHeaders[i]).ToList();
            var newVals = new decimal?[RowHeaders.Count, keep.Count];
            for (int r = 0; r < RowHeaders.Count; r++)
                for (int c = 0; c < keep.Count; c++)
                    newVals[r, c] = Values[r, keep[c]];
            ColHeaders = newCols;
            Values = newVals;
        }
    }
}
