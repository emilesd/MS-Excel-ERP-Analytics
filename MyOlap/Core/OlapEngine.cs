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

    /// <summary>
    /// Opens a model and builds the default view:
    /// Measures on rows, Time on columns, all other dimensions on POV (page filter).
    /// </summary>
    public ViewState SelectModel(long modelId)
    {
        var models = _repo.GetAllModels();
        ActiveModel = models.FirstOrDefault(m => m.Id == modelId);
        if (ActiveModel == null)
            throw new InvalidOperationException("Model not found.");

        _undo.Clear();
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
                view.PovSelections[d.Id] = roots[0].Id;
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

        for (int i = 0; i < rowMembers.Count; i++)
            rowMembers[i] = ReorderBottomUp(rowMembers[i]);

        var rowCombos = CartesianProduct(rowMembers);
        var colCombos = CartesianProduct(colMembers);

        result.RowHeaders = rowCombos;
        result.ColHeaders = colCombos;
        result.RowDimensionNames = CurrentView.RowAxes.Select(a => a.DimensionName).ToList();
        result.ColDimensionNames = CurrentView.ColAxes.Select(a => a.DimensionName).ToList();

        var dims = _repo.GetDimensions(CurrentView.ModelId);
        var dimOrder = dims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();

        result.Values = new decimal?[rowCombos.Count, colCombos.Count];

        var allFacts = _repo.GetAllFacts(CurrentView.ModelId);
        var parsedFacts = new List<(FactData fact, Dictionary<long, long> members)>();
        foreach (var f in allFacts)
        {
            var parts = f.MemberKey.Split('|');
            var memberMap = new Dictionary<long, long>();
            for (int i = 0; i < dimOrder.Count && i < parts.Length; i++)
            {
                if (long.TryParse(parts[i], out var mid) && mid > 0)
                    memberMap[dimOrder[i]] = mid;
            }
            parsedFacts.Add((f, memberMap));
        }

        var descendantCache = new Dictionary<long, Dictionary<long, int>>();
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

                    signs[child.Id] = childSign;
                    BuildSigns(child.Id, childSign);
                }
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
                        {
                            continue;
                        }
                        var signMap = descendantCache[targetMemberId];
                        if (!signMap.TryGetValue(factMemberId, out var sign))
                        {
                            match = false;
                            break;
                        }
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

        if (settings.OmitEmptyRows)
            result.OmitEmptyRows();
        if (settings.OmitEmptyColumns)
            result.OmitEmptyColumns();

        result.MemberDisplay = settings.MemberDisplay;

        return result;
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
                replacementIds = _repo.GetAllDescendants(memberId).Select(m => m.Id).ToList();
                break;
            case DrillMode.BaseOnly:
                replacementIds = _repo.GetLeafDescendants(memberId).Select(m => m.Id).ToList();
                break;
            default:
                return;
        }

        if (replacementIds.Count == 0) return;
        if (mode == DrillMode.AllGenerations || mode == DrillMode.NextGeneration) replacementIds.Insert(0, memberId);

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
            CurrentView.ColAxes.Add(rowAxis);
        }
        else if (colAxis != null)
        {
            CurrentView.ColAxes.Remove(colAxis);
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
    /// </summary>
    public void RemoveSelected(long dimensionId, long memberId)
    {
        if (CurrentView == null) return;
        PushUndo();

        foreach (var axis in CurrentView.RowAxes.Where(a => a.DimensionId == dimensionId))
            axis.VisibleMemberIds.Remove(memberId);
        foreach (var axis in CurrentView.ColAxes.Where(a => a.DimensionId == dimensionId))
            axis.VisibleMemberIds.Remove(memberId);
    }

    /// <summary>
    /// Places a picked member onto the row or column axis.
    /// </summary>
    public void PickMember(long dimensionId, long memberId, bool onRow)
    {
        if (CurrentView == null) return;
        PushUndo();

        var targetAxes = onRow ? CurrentView.RowAxes : CurrentView.ColAxes;
        var otherAxes = onRow ? CurrentView.ColAxes : CurrentView.RowAxes;

        otherAxes.RemoveAll(a => a.DimensionId == dimensionId);
        CurrentView.PovSelections.Remove(dimensionId);

        var existing = targetAxes.FirstOrDefault(a => a.DimensionId == dimensionId);
        if (existing != null)
        {
            if (!existing.VisibleMemberIds.Contains(memberId))
                existing.VisibleMemberIds.Add(memberId);
        }
        else
        {
            var dim = _repo.GetDimensions(CurrentView.ModelId)
                .FirstOrDefault(d => d.Id == dimensionId);
            targetAxes.Add(new DimensionAxis
            {
                DimensionId = dimensionId,
                DimensionName = dim?.Name ?? "Unknown",
                VisibleMemberIds = new List<long> { memberId }
            });
        }
    }

    #endregion

    #region Private helpers

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
            2 => $"{m.Name} - {m.Description}",
            _ => m.Name
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
