using System.Windows.Forms;
using MyOlap.Data;

namespace MyOlap.UI;

/// <summary>
/// Dialog for selecting an existing model or creating a new one.
/// </summary>
public class ModelBrowserForm : Form
{
    private readonly ListBox _listBox;
    private readonly Button _btnSelect;
    private readonly Button _btnNew;
    private readonly Button _btnDelete;
    private readonly Button _btnCancel;
    private List<OlapModel> _models = new();

    public long SelectedModelId { get; private set; }
    public bool CreateNew { get; private set; }

    public ModelBrowserForm()
    {
        Text = "MyOlap – Select Model";
        Width = 420;
        Height = 360;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;

        var label = new Label
        {
            Text = "Available Models:",
            Left = 12, Top = 12, Width = 380, Height = 20
        };

        _listBox = new ListBox
        {
            Left = 12, Top = 36, Width = 380, Height = 200
        };

        _btnSelect = new Button
        {
            Text = "Open", Left = 12, Top = 250, Width = 90, Height = 30,
            DialogResult = DialogResult.OK
        };
        _btnSelect.Click += (_, _) =>
        {
            if (_listBox.SelectedIndex >= 0)
                SelectedModelId = _models[_listBox.SelectedIndex].Id;
        };

        _btnNew = new Button
        {
            Text = "New Model…", Left = 110, Top = 250, Width = 100, Height = 30
        };
        _btnNew.Click += (_, _) => { CreateNew = true; DialogResult = DialogResult.OK; Close(); };

        _btnDelete = new Button
        {
            Text = "Delete", Left = 218, Top = 250, Width = 80, Height = 30
        };
        _btnDelete.Click += BtnDelete_Click;

        _btnCancel = new Button
        {
            Text = "Cancel", Left = 306, Top = 250, Width = 80, Height = 30,
            DialogResult = DialogResult.Cancel
        };

        Controls.AddRange(new Control[] { label, _listBox, _btnSelect, _btnNew, _btnDelete, _btnCancel });
        AcceptButton = _btnSelect;
        CancelButton = _btnCancel;

        LoadModels();
    }

    private void LoadModels()
    {
        _models = SqliteRepository.Instance.GetAllModels();
        _listBox.Items.Clear();
        foreach (var m in _models)
            _listBox.Items.Add($"{m.Name}   ({m.CreatedUtc:yyyy-MM-dd})");
        if (_listBox.Items.Count > 0)
            _listBox.SelectedIndex = 0;
    }

    private void BtnDelete_Click(object? sender, EventArgs e)
    {
        if (_listBox.SelectedIndex < 0) return;
        var model = _models[_listBox.SelectedIndex];
        var result = MessageBox.Show($"Delete model '{model.Name}'?", "Confirm",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (result == DialogResult.Yes)
        {
            SqliteRepository.Instance.DeleteModel(model.Id);
            LoadModels();
        }
    }
}
