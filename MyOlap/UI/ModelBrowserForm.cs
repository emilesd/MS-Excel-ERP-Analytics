using System.Windows.Forms;
using MyOlap.Data;

namespace MyOlap.UI;

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
    public long CloneFromId { get; private set; }

    public ModelBrowserForm()
    {
        AutoScaleMode = AutoScaleMode.Font;
        Text = "MyOlap \u2013 Select Model";
        Width = 680;
        Height = 400;
        MinimumSize = new System.Drawing.Size(680, 400);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var label = new Label
        {
            Text = "Available Models:",
            Left = 12, Top = 12, AutoSize = true
        };

        _listBox = new ListBox
        {
            Left = 12, Top = 44, Width = 636, Height = 224
        };

        var buttonPanel = new FlowLayoutPanel
        {
            Left = 12, Top = 276, Width = 636, Height = 44,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0)
        };

        _btnSelect = new Button
        {
            Text = "Open", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(90, 36),
            Margin = new Padding(0, 0, 8, 0),
            DialogResult = DialogResult.OK
        };
        _btnSelect.Click += (_, _) =>
        {
            if (_listBox.SelectedIndex >= 0)
                SelectedModelId = _models[_listBox.SelectedIndex].Id;
        };

        _btnNew = new Button
        {
            Text = "New Model", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(90, 36),
            Margin = new Padding(0, 0, 8, 0)
        };
        _btnNew.Click += (_, _) => { CreateNew = true; DialogResult = DialogResult.OK; Close(); };

        var btnClone = new Button
        {
            Text = "Clone", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(90, 36),
            Margin = new Padding(0, 0, 8, 0)
        };
        btnClone.Click += (_, _) =>
        {
            if (_listBox.SelectedIndex >= 0)
            {
                CloneFromId = _models[_listBox.SelectedIndex].Id;
                CreateNew = true;
                DialogResult = DialogResult.OK;
                Close();
            }
        };

        _btnDelete = new Button
        {
            Text = "Delete", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(90, 36),
            Margin = new Padding(0, 0, 8, 0)
        };
        _btnDelete.Click += BtnDelete_Click;

        _btnCancel = new Button
        {
            Text = "Cancel", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(90, 36),
            Margin = new Padding(0, 0, 0, 0),
            DialogResult = DialogResult.Cancel
        };

        buttonPanel.Controls.AddRange(new Control[] { _btnSelect, _btnNew, btnClone, _btnDelete, _btnCancel });
        Controls.AddRange(new Control[] { label, _listBox, buttonPanel });
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