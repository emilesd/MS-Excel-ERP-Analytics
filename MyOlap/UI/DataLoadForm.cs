using System.Windows.Forms;
using MyOlap.Data;

namespace MyOlap.UI;

/// <summary>
/// Dialog for loading data: pick a file, map columns to dimensions,
/// and import fact data into the model.
/// </summary>
public class DataLoadForm : Form
{
    private readonly long _modelId;
    private readonly TextBox _txtFile;
    private readonly Button _btnBrowse;
    private readonly DataGridView _dgvMapping;
    private readonly Button _btnLoad;
    private readonly Button _btnCancel;
    private readonly Label _lblStatus;
    private List<string> _headers = new();
    private List<Dimension> _dims = new();
    private string _filePath = "";

    public DataLoadForm(long modelId)
    {
        _modelId = modelId;
        AutoScaleMode = AutoScaleMode.Font;
        Text = "MyOlap \u2013 Load Data";
        Width = 860;
        Height = 560;
        MinimumSize = new System.Drawing.Size(860, 560);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblFile = new Label { Text = "Data File:", Left = 12, Top = 14, AutoSize = true };
        _txtFile = new TextBox { Left = 120, Top = 12, Width = 590, ReadOnly = true };
        _btnBrowse = new Button { Text = "\u2026", Left = 720, Top = 10, Width = 50, Height = 28 };
        _btnBrowse.Click += BtnBrowse_Click;

        var lblMap = new Label { Text = "Column \u2192 Dimension Mapping:", Left = 12, Top = 50, AutoSize = true };
        _dgvMapping = new DataGridView
        {
            Left = 12, Top = 74, Width = 810, Height = 300,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            ScrollBars = ScrollBars.Vertical
        };
        _dgvMapping.Columns.Add("Column", "File Column");
        _dgvMapping.Columns[0].Width = 220;
        _dgvMapping.Columns[0].ReadOnly = true;

        var dimCol = new DataGridViewComboBoxColumn
        {
            Name = "Dimension",
            HeaderText = "Map To",
            Width = 250
        };
        _dgvMapping.Columns.Add(dimCol);

        var typeCol = new DataGridViewComboBoxColumn
        {
            Name = "Role",
            HeaderText = "Role",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        };
        typeCol.Items.AddRange("Dimension", "Value (Numeric)", "Value (Text)", "(Skip)");
        _dgvMapping.Columns.Add(typeCol);

        _btnLoad = new Button { Text = "Load Data", Left = 570, Top = 388, Width = 140, Height = 36 };
        _btnLoad.Click += BtnLoad_Click;

        _btnCancel = new Button
        {
            Text = "Cancel", Left = 720, Top = 388, Width = 100, Height = 36,
            DialogResult = DialogResult.Cancel
        };

        _lblStatus = new Label { Text = "", Left = 12, Top = 440, AutoSize = true, MaximumSize = new System.Drawing.Size(800, 0) };

        Controls.AddRange(new Control[]
        {
            lblFile, _txtFile, _btnBrowse, lblMap, _dgvMapping,
            _btnLoad, _btnCancel, _lblStatus
        });

        _dims = SqliteRepository.Instance.GetDimensions(modelId);
    }

    private void BtnBrowse_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Filter = "Data Files|*.xlsx;*.csv;*.txt|Excel Files|*.xlsx|CSV Files|*.csv|Text Files|*.txt",
            Title = "Select Data File"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        _filePath = dlg.FileName;
        _txtFile.Text = _filePath;

        var loader = new DataLoader();
        _headers = loader.ReadHeaders(_filePath);

        _dgvMapping.Rows.Clear();
        var dimNames = _dims.Select(d => d.Name).ToList();

        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["Account"] = "Measure",
            ["Accounts"] = "Measure",
            ["Period"] = "Time",
            ["Month"] = "Time",
            ["Scenario"] = "View",
            ["Entity"] = "BU",
            ["Department"] = "BU"
        };

        foreach (var h in _headers)
        {
            var rowIdx = _dgvMapping.Rows.Add();
            var row = _dgvMapping.Rows[rowIdx];
            row.Cells["Column"].Value = h;

            var dimCell = (DataGridViewComboBoxCell)row.Cells["Dimension"];
            dimCell.Items.Clear();
            dimCell.Items.Add("(none)");
            foreach (var dn in dimNames)
                dimCell.Items.Add(dn);

            var headerNorm = h.Trim();
            if (headerNorm.Equals("Value", StringComparison.OrdinalIgnoreCase)
                || headerNorm.Equals("Amount", StringComparison.OrdinalIgnoreCase))
            {
                dimCell.Value = "(none)";
                row.Cells["Role"].Value = "Value (Numeric)";
            }
            else
            {
                var lookupName = aliases.TryGetValue(headerNorm, out var alias) ? alias : headerNorm;
                var matched = dimNames.FirstOrDefault(d => d.Equals(lookupName, StringComparison.OrdinalIgnoreCase));
                if (matched != null)
                {
                    dimCell.Value = matched;
                    row.Cells["Role"].Value = "Dimension";
                }
                else
                {
                    dimCell.Value = "(none)";
                    row.Cells["Role"].Value = "(Skip)";
                }
            }
        }
    }

    private void BtnLoad_Click(object? sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(_filePath))
        {
            MessageBox.Show("Select a data file first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        var mapping = new DataLoader.ColumnMapping();
        int valueCol = -1;
        bool valueIsText = false;

        for (int i = 0; i < _dgvMapping.Rows.Count; i++)
        {
            var role = _dgvMapping.Rows[i].Cells["Role"].Value?.ToString() ?? "(Skip)";
            var dimName = _dgvMapping.Rows[i].Cells["Dimension"].Value?.ToString() ?? "(none)";

            if (role == "Dimension" && dimName != "(none)")
            {
                var dim = _dims.FirstOrDefault(d => d.Name == dimName);
                if (dim != null)
                    mapping.ColumnToDimension[i] = dim.Id;
            }
            else if (role.StartsWith("Value"))
            {
                valueCol = i;
                valueIsText = role.Contains("Text");
            }
        }

        if (valueCol < 0)
        {
            MessageBox.Show("Assign at least one column as Value.", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        mapping.ValueColumnIndex = valueCol;
        mapping.ValueIsText = valueIsText;

        // Every model dimension must have a mapped source column (item 6b).
        var missingAlert = DataLoader.GetMissingDimensionAlert(_dims, mapping);
        if (missingAlert != null)
        {
            MessageBox.Show(missingAlert, "MyOlap", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        _lblStatus.Text = "Loading\u2026";
        Application.DoEvents();

        try
        {
            var loader = new DataLoader();
            var count = loader.LoadData(_filePath, _modelId, mapping);
            var msg = $"Successfully loaded {count:N0} records.";
            if (loader.SkippedRows > 0)
                msg += $"\n{loader.SkippedRows:N0} rows skipped (member names not found in dimensions).\nMake sure to load dimensions before loading data.";
            _lblStatus.Text = $"Loaded {count:N0} records.";
            MessageBox.Show(msg, "Done",
                MessageBoxButtons.OK, loader.SkippedRows > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
        }
        catch (Exception ex)
        {
            _lblStatus.Text = "Error loading data.";
            MessageBox.Show(ex.Message, "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}