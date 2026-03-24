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
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        Text = "MyOlap \u2013 Select Model";
        Width = 500;
        Height = 400;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var label = new Label
        {
            Text = "Available Models:",
            Left = 12, Top = 12, Width = 450, Height = 22
        };

        _listBox = new ListBox
        {
            Left = 12, Top = 38, Width = 456, Height = 230
        };

        _btnSelect = new Button
        {
            Text = "Open", Left = 12, Top = 280, Width = 100, Height = 34,
            DialogResult = DialogResult.OK
        };
        _btnSelect.Click += (_, _) =>
        {
            if (_listBox.SelectedIndex >= 0)
                SelectedModelId = _models[_listBox.SelectedIndex].Id;
        };

        _btnNew = new Button
        {
            Text = "New Model\u2026", Left = 122, Top = 280, Width = 120, Height = 34
        };
        _btnNew.Click += (_, _) => { CreateNew = true; DialogResult = DialogResult.OK; Close(); };

        _btnDelete = new Button
        {
            Text = "Delete", Left = 252, Top = 280, Width = 100, Height = 34
        };
        _btnDelete.Click += BtnDelete_Click;

        _btnCancel = new Button
        {
            Text = "Cancel", Left = 362, Top = 280, Width = 100, Height = 34,
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
