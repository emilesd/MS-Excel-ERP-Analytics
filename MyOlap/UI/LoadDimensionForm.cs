using System.Windows.Forms;
using System.Globalization;
using System.IO;
using MyOlap.Core;
using MyOlap.Data;
using OfficeOpenXml;
using CsvHelper;
using CsvHelper.Configuration;

namespace MyOlap.UI;

public class LoadDimensionForm : Form
{
    private readonly long _modelId;
    private long _preSelectedDimId;
    private readonly TextBox _txtFile;
    private readonly ComboBox _cbSheet;
    private readonly ComboBox _cbParent;
    private readonly ComboBox _cbChild;
    private readonly ComboBox _cbDescription;
    private readonly ComboBox _cbConsolOp;
    private readonly ComboBox _cbFormula;
    private readonly ComboBox _cbTimeBalance;
    private readonly ComboBox _cbDimension;
    private readonly Label _lblStatus;
    private List<Dimension> _dims = new();
    private string _filePath = "";
    private bool _isCsv;
    private List<string> _sheetNames = new();
    private List<string> _columnHeaders = new();

    public LoadDimensionForm(long modelId, long preSelectedDimId = 0)
    {
        _modelId = modelId;
        AutoScaleMode = AutoScaleMode.Font;
        Text = "MyOlap \u2013 Load Dimension";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Padding = new Padding(24, 20, 24, 20);

        int lx = 28;
        int cw = 580;
        int y = 28;
        int lblGap = 30;
        int rowGap = 44;

        var lblFile = new Label { Text = "Source File:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _txtFile = new TextBox { Left = lx, Top = y, Width = cw - 60, ReadOnly = true };
        var btnBrowse = new Button { Text = "\u2026", Left = lx + cw - 54, Top = y - 2, Width = 54, Height = 28 };
        btnBrowse.Click += BtnBrowse_Click;

        y += rowGap;
        var lblSheet = new Label { Text = "Sheet / Tab:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbSheet = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };
        _cbSheet.SelectedIndexChanged += CbSheet_Changed;

        y += rowGap;
        var sep = new Label { Text = "Column Mapping:", Left = lx, Top = y, AutoSize = true, Font = new System.Drawing.Font(System.Drawing.SystemFonts.DefaultFont, System.Drawing.FontStyle.Bold) };

        y += lblGap + 6;
        var lblParent = new Label { Text = "Parent Column:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbParent = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap;
        var lblChild = new Label { Text = "Child Column:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbChild = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap;
        var lblDesc = new Label { Text = "Description:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbDescription = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap;
        var lblOp = new Label { Text = "Consol Operator Column:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbConsolOp = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap;
        var lblFormula = new Label { Text = "Formula Column:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbFormula = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap;
        var lblTimeBal = new Label { Text = "Time Balance Column:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbTimeBalance = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap;
        var lblDim = new Label { Text = "Target Dimension:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbDimension = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap + 20;
        int btnLoadW = 80;
        int btnAllW = 240;
        int btnCloseW = 80;
        int btnAllLeft = lx + (cw - btnAllW) / 2;

        var btnLoad = new Button { Text = "Load", Left = lx, Top = y, Width = btnLoadW, Height = 34 };
        btnLoad.Click += BtnLoad_Click;

        var btnLoadAll = new Button { Text = "All Dimensions", Left = btnAllLeft, Top = y, Width = btnAllW, Height = 34 };
        btnLoadAll.Click += BtnLoadAll_Click;

        var btnClose = new Button
        {
            Text = "Close", Left = lx + cw - btnCloseW, Top = y, Width = btnCloseW, Height = 34,
            DialogResult = DialogResult.OK
        };

        y += 44;
        _lblStatus = new Label { Text = "", Left = lx, Top = y, AutoSize = true, MaximumSize = new System.Drawing.Size(cw, 0) };

        y += 80;
        Width = lx + cw + 44;
        Height = (int)(y * 1.12);
        MinimumSize = new System.Drawing.Size(Width, Height);

        Controls.AddRange(new Control[]
        {
            lblFile, _txtFile, btnBrowse,
            lblSheet, _cbSheet,
            sep,
            lblParent, _cbParent,
            lblChild, _cbChild,
            lblDesc, _cbDescription,
            lblOp, _cbConsolOp,
            lblFormula, _cbFormula,
            lblTimeBal, _cbTimeBalance,
            lblDim, _cbDimension,
            btnLoad, btnLoadAll, btnClose, _lblStatus
        });

        AcceptButton = btnLoad;
        _preSelectedDimId = preSelectedDimId;
        _dims = SqliteRepository.Instance.GetDimensions(modelId);
        foreach (var d in _dims)
            _cbDimension.Items.Add(d.Name);
        if (_cbDimension.Items.Count > 0)
        {
            int preIdx = preSelectedDimId > 0 ? _dims.FindIndex(d => d.Id == preSelectedDimId) : -1;
            _cbDimension.SelectedIndex = preIdx >= 0 ? preIdx : 0;
        }
    }

    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Excel Files|*.xlsx;*.xls|CSV Files|*.csv|All Files|*.*",
            Title = "Select Dimension File"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _filePath = dlg.FileName;
        _txtFile.Text = _filePath;
        _isCsv = _filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);

        try
        {
            _sheetNames.Clear();
            _cbSheet.Items.Clear();

            if (_isCsv)
            {
                var name = Path.GetFileNameWithoutExtension(_filePath);
                _sheetNames.Add(name);
                _cbSheet.Items.Add(name);
            }
            else
            {
                ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
                using var pkg = new ExcelPackage(new FileInfo(_filePath));
                foreach (var ws in pkg.Workbook.Worksheets)
                {
                    _sheetNames.Add(ws.Name);
                    _cbSheet.Items.Add(ws.Name);
                }
            }

            if (_cbSheet.Items.Count > 0)
            {
                // If opened with a pre-selected dim, auto-select the matching sheet so it doesn't
                // immediately flip back to the first sheet's dimension when AutoMapColumns runs.
                int sheetIdx = 0;
                if (_preSelectedDimId > 0)
                {
                    var dimName = _dims.FirstOrDefault(d => d.Id == _preSelectedDimId)?.Name ?? "";
                    int match = _sheetNames.FindIndex(s => s.Equals(dimName, StringComparison.OrdinalIgnoreCase));
                    if (match >= 0) sheetIdx = match;
                }
                _cbSheet.SelectedIndex = sheetIdx;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading file: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void CbSheet_Changed(object? sender, EventArgs e)
    {
        if (_cbSheet.SelectedIndex < 0 || string.IsNullOrEmpty(_filePath)) return;

        try
        {
            _columnHeaders.Clear();
            _cbParent.Items.Clear();
            _cbChild.Items.Clear();
            _cbDescription.Items.Clear();
            _cbDescription.Items.Add("(none)");
            _cbConsolOp.Items.Clear();
            _cbConsolOp.Items.Add("(none)");
            _cbFormula.Items.Clear();
            _cbFormula.Items.Add("(none)");
            _cbTimeBalance.Items.Clear();
            _cbTimeBalance.Items.Add("(none)");

            if (_isCsv)
            {
                using var reader = new StreamReader(_filePath);
                using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true });
                csv.Read();
                csv.ReadHeader();
                if (csv.HeaderRecord != null)
                {
                    foreach (var header in csv.HeaderRecord)
                    {
                        var h = string.IsNullOrEmpty(header.Trim()) ? $"Column {_columnHeaders.Count + 1}" : header.Trim();
                        _columnHeaders.Add(h);
                        _cbParent.Items.Add(h);
                        _cbChild.Items.Add(h);
                        _cbDescription.Items.Add(h);
                        _cbConsolOp.Items.Add(h);
                        _cbFormula.Items.Add(h);
                        _cbTimeBalance.Items.Add(h);
                    }
                }
            }
            else
            {
                ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
                using var pkg = new ExcelPackage(new FileInfo(_filePath));
                var ws = pkg.Workbook.Worksheets[_cbSheet.SelectedIndex];
                if (ws.Dimension == null) return;

                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    var header = ws.Cells[1, col].Text.Trim();
                    if (string.IsNullOrEmpty(header)) header = $"Column {col}";
                    _columnHeaders.Add(header);
                    _cbParent.Items.Add(header);
                    _cbChild.Items.Add(header);
                    _cbDescription.Items.Add(header);
                    _cbConsolOp.Items.Add(header);
                    _cbFormula.Items.Add(header);
                    _cbTimeBalance.Items.Add(header);
                }
            }

            AutoMapColumns();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error reading sheet: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void AutoMapColumns()
    {
        var parentAliases = new[] { "parent", "parent_id", "parentid", "parent member", "parentmember" };
        var childAliases = new[] { "child", "child_id", "childid", "member", "child member", "childmember", "name" };
        var descAliases = new[] { "description", "desc", "alias", "label" };
        var opAliases = new[] { "consolop", "consol", "operator", "op", "sign" };
        var formulaAliases = new[] { "formula", "calc", "calculation" };
        var timeBalAliases = new[] { "timebalance", "time_balance", "timebal", "timbal" };

        for (int i = 0; i < _columnHeaders.Count; i++)
        {
            var h = _columnHeaders[i].ToLowerInvariant();
            if (parentAliases.Any(a => h.Contains(a)) && _cbParent.SelectedIndex < 0)
                _cbParent.SelectedIndex = i;
            if (childAliases.Any(a => h.Contains(a)) && _cbChild.SelectedIndex < 0)
                _cbChild.SelectedIndex = i;
            if (descAliases.Any(a => h.Contains(a)) && _cbDescription.SelectedIndex <= 0)
                _cbDescription.SelectedIndex = i + 1;
            if (opAliases.Any(a => h.Contains(a)) && _cbConsolOp.SelectedIndex <= 0)
                _cbConsolOp.SelectedIndex = i + 1;
            if (formulaAliases.Any(a => h.Contains(a)) && _cbFormula.SelectedIndex <= 0)
                _cbFormula.SelectedIndex = i + 1;
            if (timeBalAliases.Any(a => h.Contains(a)) && _cbTimeBalance.SelectedIndex <= 0)
                _cbTimeBalance.SelectedIndex = i + 1;
        }

        if (_cbParent.SelectedIndex < 0 && _cbParent.Items.Count >= 1) _cbParent.SelectedIndex = 0;
        if (_cbChild.SelectedIndex < 0 && _cbChild.Items.Count >= 2) _cbChild.SelectedIndex = 1;
        if (_cbDescription.SelectedIndex < 0) _cbDescription.SelectedIndex = 0;
        if (_cbConsolOp.SelectedIndex < 0) _cbConsolOp.SelectedIndex = 0;
        if (_cbFormula.SelectedIndex < 0) _cbFormula.SelectedIndex = 0;
        if (_cbTimeBalance.SelectedIndex < 0) _cbTimeBalance.SelectedIndex = 0;

        var sheetName = _cbSheet.Text.ToLowerInvariant();
        for (int i = 0; i < _dims.Count; i++)
        {
            if (_dims[i].Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
            {
                _cbDimension.SelectedIndex = i;
                break;
            }
        }
    }

    private List<(string parent, string child, string desc, string consolOp, string formula, string timeBalance)> ReadRows()
    {
        int parentCol = _cbParent.SelectedIndex;
        int childCol = _cbChild.SelectedIndex;
        int descCol = _cbDescription.SelectedIndex - 1;
        int opCol = _cbConsolOp.SelectedIndex - 1;
        int formulaCol = _cbFormula.SelectedIndex - 1;
        int timeBalCol = _cbTimeBalance.SelectedIndex - 1;
        var rows = new List<(string parent, string child, string desc, string consolOp, string formula, string timeBalance)>();

        if (_isCsv)
        {
            using var reader = new StreamReader(_filePath);
            using var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true });
            while (csv.Read())
            {
                var parent = csv.GetField(parentCol)?.Trim() ?? "";
                var child = csv.GetField(childCol)?.Trim() ?? "";
                var desc = descCol >= 0 ? (csv.GetField(descCol)?.Trim() ?? "") : "";
                var consolOp = opCol >= 0 ? (csv.GetField(opCol)?.Trim() ?? "+") : "+";
                var formula = formulaCol >= 0 ? (csv.GetField(formulaCol)?.Trim() ?? "") : "";
                var timeBalance = timeBalCol >= 0 ? (csv.GetField(timeBalCol)?.Trim() ?? "") : "";
                if (!string.IsNullOrEmpty(child) && !parent.Equals(child, StringComparison.OrdinalIgnoreCase))
                    rows.Add((parent, child, desc, consolOp, formula, timeBalance));
            }
        }
        else
        {
            ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
            using var pkg = new ExcelPackage(new FileInfo(_filePath));
            var ws = pkg.Workbook.Worksheets[_cbSheet.SelectedIndex];
            if (ws.Dimension == null) return rows;

            for (int r = 2; r <= ws.Dimension.End.Row; r++)
            {
                var parent = ws.Cells[r, parentCol + 1].Text.Trim();
                var child = ws.Cells[r, childCol + 1].Text.Trim();
                var desc = descCol >= 0 ? ws.Cells[r, descCol + 1].Text.Trim() : "";
                var consolOp = opCol >= 0 ? ws.Cells[r, opCol + 1].Text.Trim() : "+";
                var formula = formulaCol >= 0 ? ws.Cells[r, formulaCol + 1].Text.Trim() : "";
                var timeBalance = timeBalCol >= 0 ? ws.Cells[r, timeBalCol + 1].Text.Trim() : "";
                if (!string.IsNullOrEmpty(child) && !parent.Equals(child, StringComparison.OrdinalIgnoreCase))
                    rows.Add((parent, child, desc, consolOp, formula, timeBalance));
            }
        }

        return rows;
    }

    private void BtnLoad_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_filePath))
        {
            MessageBox.Show("Select a file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (_cbParent.SelectedIndex < 0 || _cbChild.SelectedIndex < 0)
        {
            MessageBox.Show("Map Parent and Child columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (_cbDimension.SelectedIndex < 0)
        {
            MessageBox.Show("Select a target dimension.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var dim = _dims[_cbDimension.SelectedIndex];
        _lblStatus.Text = "Loading\u2026";
        Application.DoEvents();

        try
        {
            var rows = ReadRows();
            int count = LoadDimensionFromRows(dim, rows);
            _lblStatus.Text = $"Done: {count} members processed.";
            MessageBox.Show($"Dimension '{dim.Name}': {count} members processed.\nFact data has been preserved.",
                "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _lblStatus.Text = "Error.";
            MessageBox.Show($"Error loading dimension: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void BtnLoadAll_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_filePath))
        {
            MessageBox.Show("Select a file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        if (_cbParent.SelectedIndex < 0 || _cbChild.SelectedIndex < 0)
        {
            MessageBox.Show("Map Parent and Child columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        _lblStatus.Text = "Loading all dimensions\u2026";
        Application.DoEvents();

        try
        {
            var results = new List<string>();
            var skipped = new List<string>();
            int originalDimIndex = _cbDimension.SelectedIndex;

            // Creating a brand-new dimension is only allowed on an empty model (no fact data).
            bool needsNewDimension = _sheetNames.Any(sheetName =>
                !_dims.Any(d => d.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase)));
            if (needsNewDimension)
            {
                var mgr = new ModelManager();
                if (!mgr.CanAddNewDimension(_modelId, out var err))
                {
                    _lblStatus.Text = "Blocked.";
                    MessageBox.Show(err, "MyOlap", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // Iterate over every sheet in the file; find or create a matching dimension.
            for (int s = 0; s < _sheetNames.Count; s++)
            {
                var sheetName = _sheetNames[s];

                // Find existing dimension whose name matches the sheet name
                int dimIdx = _dims.FindIndex(d => d.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

                if (dimIdx < 0)
                {
                    // No dimension with this name — create one (model already verified empty above).
                    var newDimId = SqliteRepository.Instance.InsertDimension(new Dimension
                    {
                        ModelId = _modelId,
                        Name = sheetName,
                        DimType = DimensionType.UserDefined,
                        SortOrder = _dims.Count
                    });
                    var newDim = new Dimension
                    {
                        Id = newDimId, ModelId = _modelId, Name = sheetName,
                        DimType = DimensionType.UserDefined, SortOrder = _dims.Count
                    };
                    _dims.Add(newDim);
                    _cbDimension.Items.Add(sheetName);
                    dimIdx = _dims.Count - 1;
                }

                _cbSheet.SelectedIndex = s;
                _cbDimension.SelectedIndex = dimIdx;
                Application.DoEvents();

                var rows = ReadRows();
                if (rows.Count == 0) { skipped.Add($"{sheetName} (no data rows)"); continue; }

                int count = LoadDimensionFromRows(_dims[dimIdx], rows);
                results.Add($"{_dims[dimIdx].Name} ({count} members)");
            }

            if (originalDimIndex >= 0 && originalDimIndex < _cbDimension.Items.Count)
                _cbDimension.SelectedIndex = originalDimIndex;

            if (results.Count == 0)
            {
                _lblStatus.Text = "No sheets with data found.";
                MessageBox.Show("No sheets in the file contained data rows.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                _lblStatus.Text = $"Loaded {results.Count} dimensions.";
                var msg = $"Loaded {results.Count} dimensions:\n\n{string.Join("\n", results)}";
                if (skipped.Count > 0)
                    msg += $"\n\nSkipped (empty):\n{string.Join("\n", skipped)}";
                MessageBox.Show(msg, "All Dimensions Loaded", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            _lblStatus.Text = "Error.";
            MessageBox.Show($"Error loading dimensions: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private int LoadDimensionFromRows(Dimension dim, List<(string parent, string child, string desc, string consolOp, string formula, string timeBalance)> rows)
    {
        var repo = SqliteRepository.Instance;
        var existing = repo.GetMembersByNameForDimension(dim.Id);
        var memberMap = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in existing)
            memberMap[kvp.Key] = kvp.Value.Id;

        // Pre-scan rows in file order: the parent column of the FIRST row where each child name
        // appears is its canonical (base) parent. All later rows with the same child name are shared.
        var firstOccurrenceParent = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (p, c, _, _, _, _) in rows)
        {
            if (!firstOccurrenceParent.ContainsKey(c))
                firstOccurrenceParent[c] = p;
        }

        var allChildrenSet = new HashSet<string>(rows.Select(r => r.child), StringComparer.OrdinalIgnoreCase);
        var rootNames = rows
            .Select(r => r.parent)
            .Where(p => !string.IsNullOrEmpty(p) && !allChildrenSet.Contains(p))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        int order = 0;
        int total = 0;

        foreach (var rootName in rootNames)
        {
            if (existing.TryGetValue(rootName, out var ex))
            {
                ex.ParentId = null; ex.Description = rootName; ex.Level = 0; ex.SortOrder = order++;
                repo.UpdateMember(ex);
                memberMap[rootName] = ex.Id;
            }
            else if (!memberMap.ContainsKey(rootName))
            {
                var id = repo.InsertMember(new Member { DimensionId = dim.Id, Name = rootName, Description = rootName, Level = 0, SortOrder = order++ });
                memberMap[rootName] = id;
            }
            total++;
        }

        var processed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var queue = new Queue<string>(rootNames);
        // Deferred shared insertions: non-canonical occurrences encountered before their base was inserted.
        var deferred = new List<(string parentName, string childName, string desc, string op, string formula, string timeBal)>();

        while (queue.Count > 0)
        {
            var parentName = queue.Dequeue();
            if (!memberMap.TryGetValue(parentName, out var parentId)) continue;

            var children = rows.Where(r => r.parent.Equals(parentName, StringComparison.OrdinalIgnoreCase)).ToList();
            foreach (var (_, childName, desc, op, formula, timeBal) in children)
            {
                if (childName.Equals(parentName, StringComparison.OrdinalIgnoreCase)) continue;
                var resolvedDesc = string.IsNullOrEmpty(desc) ? childName : desc;
                var resolvedOp = NormalizeOp(op);
                var parentMember = repo.GetMember(parentId);
                int level = (parentMember?.Level ?? 0) + 1;

                bool isCanonical = firstOccurrenceParent.TryGetValue(childName, out var canonParent)
                                   && canonParent.Equals(parentName, StringComparison.OrdinalIgnoreCase);

                if (existing.TryGetValue(childName, out var exChild))
                {
                    // Member already in DB from a previous load — update or create shared copy
                    if (exChild.ParentId.HasValue && exChild.ParentId.Value != parentId && memberMap.ContainsKey(childName))
                    {
                        var sharedName = $"{childName}__shared_{parentName}";
                        if (!existing.ContainsKey(sharedName) && !memberMap.ContainsKey(sharedName))
                        {
                            var sharedId = repo.InsertMember(new Member
                            {
                                DimensionId = dim.Id, ParentId = parentId, Name = sharedName,
                                Description = resolvedDesc, Level = level, SortOrder = order++,
                                ConsolOperator = resolvedOp, SharedFromId = exChild.Id
                            });
                            memberMap[sharedName] = sharedId;
                        }
                    }
                    else
                    {
                        exChild.ParentId = parentId; exChild.Description = resolvedDesc;
                        exChild.Level = level; exChild.SortOrder = order++; exChild.ConsolOperator = resolvedOp;
                        exChild.Formula = formula; exChild.TimeBalance = timeBal;
                        repo.UpdateMember(exChild);
                        memberMap[childName] = exChild.Id;
                    }
                }
                else if (memberMap.ContainsKey(childName))
                {
                    // Base already inserted earlier in this session — create shared copy
                    var baseId = memberMap[childName];
                    var sharedName = $"{childName}__shared_{parentName}";
                    if (!memberMap.ContainsKey(sharedName))
                    {
                        var sharedId = repo.InsertMember(new Member
                        {
                            DimensionId = dim.Id, ParentId = parentId, Name = sharedName,
                            Description = resolvedDesc, Level = level, SortOrder = order++,
                            ConsolOperator = resolvedOp, SharedFromId = baseId
                        });
                        memberMap[sharedName] = sharedId;
                    }
                }
                else if (isCanonical)
                {
                    // First occurrence in file — insert as base member
                    var id = repo.InsertMember(new Member
                    {
                        DimensionId = dim.Id, ParentId = parentId, Name = childName,
                        Description = resolvedDesc, Level = level, SortOrder = order++,
                        ConsolOperator = resolvedOp, Formula = formula, TimeBalance = timeBal
                    });
                    memberMap[childName] = id;
                }
                else
                {
                    // Non-canonical occurrence but base not yet in memberMap (BFS reached this parent
                    // before the canonical parent) — defer until base is available.
                    deferred.Add((parentName, childName, resolvedDesc, resolvedOp, formula, timeBal));
                }
                total++;

                if (!processed.Contains(childName)) { processed.Add(childName); queue.Enqueue(childName); }
            }
        }

        // Insert deferred shared members now that all base members have been created.
        foreach (var (parentName, childName, resolvedDesc, resolvedOp, formula, timeBal) in deferred)
        {
            if (!memberMap.TryGetValue(parentName, out var parentId)) continue;
            var parentMember = repo.GetMember(parentId);
            int level = (parentMember?.Level ?? 0) + 1;

            if (memberMap.TryGetValue(childName, out var baseId))
            {
                var sharedName = $"{childName}__shared_{parentName}";
                if (!memberMap.ContainsKey(sharedName))
                {
                    var sharedId = repo.InsertMember(new Member
                    {
                        DimensionId = dim.Id, ParentId = parentId, Name = sharedName,
                        Description = resolvedDesc, Level = level, SortOrder = order++,
                        ConsolOperator = resolvedOp, SharedFromId = baseId
                    });
                    memberMap[sharedName] = sharedId;
                }
            }
            else
            {
                // Base still not found — insert as base (edge case fallback)
                var id = repo.InsertMember(new Member
                {
                    DimensionId = dim.Id, ParentId = parentId, Name = childName,
                    Description = resolvedDesc, Level = level, SortOrder = order++,
                    ConsolOperator = resolvedOp, Formula = formula, TimeBalance = timeBal
                });
                memberMap[childName] = id;
            }
        }

        return total;
    }

    private static string NormalizeOp(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "+";
        raw = raw.Trim();
        if (raw == "-" || raw.Equals("subtract", StringComparison.OrdinalIgnoreCase)) return "-";
        if (raw == "x" || raw == "X" || raw.Equals("exclude", StringComparison.OrdinalIgnoreCase)) return "x";
        if (raw.Equals("ignore", StringComparison.OrdinalIgnoreCase)) return "+";
        return "+";
    }
}
