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
        Text = "MyOlap – Load Data";
        Width = 600;
        Height = 480;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblFile = new Label { Text = "Data File:", Left = 12, Top = 14, Width = 70 };
        _txtFile = new TextBox { Left = 85, Top = 12, Width = 400, ReadOnly = true };
        _btnBrowse = new Button { Text = "…", Left = 490, Top = 11, Width = 40, Height = 24 };
        _btnBrowse.Click += BtnBrowse_Click;

        var lblMap = new Label { Text = "Column → Dimension Mapping:", Left = 12, Top = 48, Width = 300 };
        _dgvMapping = new DataGridView
        {
            Left = 12, Top = 70, Width = 560, Height = 280,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect
        };
        _dgvMapping.Columns.Add("Column", "File Column");
        _dgvMapping.Columns[0].Width = 180;
        _dgvMapping.Columns[0].ReadOnly = true;

        var dimCol = new DataGridViewComboBoxColumn
        {
            Name = "Dimension",
            HeaderText = "Map To Dimension",
            Width = 200
        };
        _dgvMapping.Columns.Add(dimCol);

        var typeCol = new DataGridViewComboBoxColumn
        {
            Name = "Role",
            HeaderText = "Role",
            Width = 120
        };
        typeCol.Items.AddRange("Dimension", "Value (Numeric)", "Value (Text)", "(Skip)");
        _dgvMapping.Columns.Add(typeCol);

        _btnLoad = new Button { Text = "Load Data", Left = 380, Top = 365, Width = 100, Height = 30 };
        _btnLoad.Click += BtnLoad_Click;

        _btnCancel = new Button
        {
            Text = "Cancel", Left = 490, Top = 365, Width = 80, Height = 30,
            DialogResult = DialogResult.Cancel
        };

        _lblStatus = new Label { Text = "", Left = 12, Top = 405, Width = 560, Height = 24 };

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
            dimCell.Value = "(none)";

            row.Cells["Role"].Value = "(Skip)";
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

        _lblStatus.Text = "Loading…";
        Application.DoEvents();

        try
        {
            var loader = new DataLoader();
            var count = loader.LoadData(_filePath, _modelId, mapping);
            _lblStatus.Text = $"Loaded {count:N0} records.";
            MessageBox.Show($"Successfully loaded {count:N0} records.", "Done",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _lblStatus.Text = "Error loading data.";
            MessageBox.Show(ex.Message, "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
