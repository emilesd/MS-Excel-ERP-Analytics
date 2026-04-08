using System.Windows.Forms;
using System.Globalization;
using System.IO;
using MyOlap.Data;
using OfficeOpenXml;
using CsvHelper;
using CsvHelper.Configuration;

namespace MyOlap.UI;

public class LoadDimensionForm : Form
{
    private readonly long _modelId;
    private readonly TextBox _txtFile;
    private readonly ComboBox _cbSheet;
    private readonly ComboBox _cbParent;
    private readonly ComboBox _cbChild;
    private readonly ComboBox _cbDescription;
    private readonly ComboBox _cbConsolOp;
    private readonly ComboBox _cbDimension;
    private readonly Label _lblStatus;
    private List<Dimension> _dims = new();
    private string _filePath = "";
    private bool _isCsv;
    private List<string> _sheetNames = new();
    private List<string> _columnHeaders = new();

    public LoadDimensionForm(long modelId)
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
        int cw = 420;
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
        var lblDim = new Label { Text = "Target Dimension:", Left = lx, Top = y, AutoSize = true };
        y += lblGap;
        _cbDimension = new ComboBox { Left = lx, Top = y, Width = cw, DropDownStyle = ComboBoxStyle.DropDownList };

        y += rowGap + 20;
        var btnLoad = new Button { Text = "Load", Left = lx + cw - 220, Top = y, Width = 110, Height = 34 };
        btnLoad.Click += BtnLoad_Click;

        var btnCancel = new Button
        {
            Text = "Cancel", Left = lx + cw - 100, Top = y, Width = 100, Height = 34,
            DialogResult = DialogResult.Cancel
        };

        _lblStatus = new Label { Text = "", Left = lx, Top = y + 8, AutoSize = true, MaximumSize = new System.Drawing.Size(220, 0) };

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
            lblDim, _cbDimension,
            btnLoad, btnCancel, _lblStatus
        });

        CancelButton = btnCancel;
        _dims = SqliteRepository.Instance.GetDimensions(modelId);
        foreach (var d in _dims)
            _cbDimension.Items.Add(d.Name);
        if (_cbDimension.Items.Count > 0)
            _cbDimension.SelectedIndex = 0;
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
                _cbSheet.SelectedIndex = 0;
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
        }

        if (_cbParent.SelectedIndex < 0 && _cbParent.Items.Count >= 1) _cbParent.SelectedIndex = 0;
        if (_cbChild.SelectedIndex < 0 && _cbChild.Items.Count >= 2) _cbChild.SelectedIndex = 1;
        if (_cbDescription.SelectedIndex < 0) _cbDescription.SelectedIndex = 0;
        if (_cbConsolOp.SelectedIndex < 0) _cbConsolOp.SelectedIndex = 0;

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

    private List<(string parent, string child, string desc, string consolOp)> ReadRows()
    {
        int parentCol = _cbParent.SelectedIndex;
        int childCol = _cbChild.SelectedIndex;
        int descCol = _cbDescription.SelectedIndex - 1;
        int opCol = _cbConsolOp.SelectedIndex - 1;
        var rows = new List<(string parent, string child, string desc, string consolOp)>();

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
                if (!string.IsNullOrEmpty(child))
                    rows.Add((parent, child, desc, consolOp));
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
                if (!string.IsNullOrEmpty(child))
                    rows.Add((parent, child, desc, consolOp));
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
            var repo = SqliteRepository.Instance;

            repo.ClearDimensionMembers(dim.Id);

            var memberMap = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

            var allChildren = new HashSet<string>(rows.Select(r => r.child), StringComparer.OrdinalIgnoreCase);
            var allParents = new HashSet<string>(rows.Select(r => r.parent).Where(p => !string.IsNullOrEmpty(p)), StringComparer.OrdinalIgnoreCase);
            var rootNames = allParents.Except(allChildren, StringComparer.OrdinalIgnoreCase).ToList();

            int order = 0;

            foreach (var rootName in rootNames)
            {
                if (!memberMap.ContainsKey(rootName))
                {
                    var id = repo.InsertMember(new Member
                    {
                        DimensionId = dim.Id,
                        Name = rootName,
                        Description = rootName,
                        Level = 0,
                        SortOrder = order++
                    });
                    memberMap[rootName] = id;
                }
            }

            int loaded = 0;
            var processed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var queue = new Queue<string>(rootNames);

            while (queue.Count > 0)
            {
                var parentName = queue.Dequeue();
                if (!memberMap.TryGetValue(parentName, out var parentId)) continue;

                var children = rows.Where(r => r.parent.Equals(parentName, StringComparison.OrdinalIgnoreCase)).ToList();
                foreach (var (_, childName, desc, op) in children)
                {
                    if (!memberMap.ContainsKey(childName))
                    {
                        var parentMember = repo.GetMember(parentId);
                        var id = repo.InsertMember(new Member
                        {
                            DimensionId = dim.Id,
                            ParentId = parentId,
                            Name = childName,
                            Description = string.IsNullOrEmpty(desc) ? childName : desc,
                            Level = (parentMember?.Level ?? 0) + 1,
                            SortOrder = order++,
                            ConsolOperator = NormalizeOp(op)
                        });
                        memberMap[childName] = id;
                        loaded++;
                    }

                    if (!processed.Contains(childName))
                    {
                        processed.Add(childName);
                        queue.Enqueue(childName);
                    }
                }
            }

            foreach (var row in rows)
            {
                if (!memberMap.ContainsKey(row.child))
                {
                    long? pid = string.IsNullOrEmpty(row.parent) ? null : 
                        memberMap.TryGetValue(row.parent, out var pv) ? pv : null;
                    var parentMember = pid.HasValue ? repo.GetMember(pid.Value) : null;
                    var id = repo.InsertMember(new Member
                    {
                        DimensionId = dim.Id,
                        ParentId = pid,
                        Name = row.child,
                        Description = string.IsNullOrEmpty(row.desc) ? row.child : row.desc,
                        Level = (parentMember?.Level ?? 0) + 1,
                        SortOrder = order++,
                        ConsolOperator = NormalizeOp(row.consolOp)
                    });
                    memberMap[row.child] = id;
                    loaded++;
                }
            }

            _lblStatus.Text = $"Loaded {loaded} members.";
            MessageBox.Show($"Successfully loaded {loaded} members into '{dim.Name}'.",
                "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
        }
        catch (Exception ex)
        {
            _lblStatus.Text = "Error.";
            MessageBox.Show($"Error loading dimension: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
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