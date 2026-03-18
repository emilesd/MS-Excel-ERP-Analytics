using System.Windows.Forms;
using MyOlap.Data;

namespace MyOlap.UI;

/// <summary>
/// Settings dialog for retrieval options:
/// - Omit Empty Rows
/// - Omit Empty Columns
/// - Show Member Name, Description, or Both
/// </summary>
public class SettingsForm : Form
{
    private readonly CheckBox _chkOmitRows;
    private readonly CheckBox _chkOmitCols;
    private readonly ComboBox _cbDisplay;
    private readonly Button _btnOk;
    private readonly Button _btnCancel;

    public ModelSettings Settings { get; private set; }

    public SettingsForm(ModelSettings current)
    {
        Settings = current;
        Text = "MyOlap – Settings";
        Width = 360;
        Height = 260;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;

        _chkOmitRows = new CheckBox
        {
            Text = "Omit Empty Rows",
            Left = 20, Top = 20, Width = 300,
            Checked = current.OmitEmptyRows
        };

        _chkOmitCols = new CheckBox
        {
            Text = "Omit Empty Columns",
            Left = 20, Top = 50, Width = 300,
            Checked = current.OmitEmptyColumns
        };

        var lblDisplay = new Label { Text = "Member Display:", Left = 20, Top = 86, Width = 120 };
        _cbDisplay = new ComboBox
        {
            Left = 145, Top = 84, Width = 180,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cbDisplay.Items.AddRange(new object[] { "Name Only", "Description Only", "Name and Description" });
        _cbDisplay.SelectedIndex = current.MemberDisplay;

        _btnOk = new Button
        {
            Text = "OK", Left = 150, Top = 140, Width = 80, Height = 30,
            DialogResult = DialogResult.OK
        };
        _btnOk.Click += (_, _) =>
        {
            Settings = new ModelSettings
            {
                ModelId = current.ModelId,
                OmitEmptyRows = _chkOmitRows.Checked,
                OmitEmptyColumns = _chkOmitCols.Checked,
                MemberDisplay = _cbDisplay.SelectedIndex
            };
        };

        _btnCancel = new Button
        {
            Text = "Cancel", Left = 240, Top = 140, Width = 80, Height = 30,
            DialogResult = DialogResult.Cancel
        };

        Controls.AddRange(new Control[] { _chkOmitRows, _chkOmitCols, lblDisplay, _cbDisplay, _btnOk, _btnCancel });
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;
    }
}
