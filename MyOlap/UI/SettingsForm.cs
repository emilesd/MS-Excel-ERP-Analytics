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
        AutoScaleMode = AutoScaleMode.Font;
        Text = "MyOlap \u2013 Settings";
        Width = 420;
        Height = 300;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Padding = new Padding(24, 20, 24, 16);

        _chkOmitRows = new CheckBox
        {
            Text = "Omit Empty Rows", Left = 24, Top = 24, AutoSize = true,
            Checked = current.OmitEmptyRows
        };

        _chkOmitCols = new CheckBox
        {
            Text = "Omit Empty Columns", Left = 24, Top = 58, AutoSize = true,
            Checked = current.OmitEmptyColumns
        };

        var lblDisplay = new Label { Text = "Member Display:", Left = 24, Top = 100, AutoSize = true };
        _cbDisplay = new ComboBox
        {
            Left = 24, Top = 130, Width = 340,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cbDisplay.Items.AddRange(new object[] { "Name Only", "Description Only", "Name and Description" });
        _cbDisplay.SelectedIndex = current.MemberDisplay;

        _btnOk = new Button
        {
            Text = "OK", Width = 100, Height = 34,
            DialogResult = DialogResult.OK
        };
        _btnOk.Left = 140;
        _btnOk.Top = 182;
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
            Text = "Cancel", Width = 100, Height = 34,
            DialogResult = DialogResult.Cancel
        };
        _btnCancel.Left = 252;
        _btnCancel.Top = 182;

        Controls.AddRange(new Control[] { _chkOmitRows, _chkOmitCols, lblDisplay, _cbDisplay, _btnOk, _btnCancel });
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;
    }
}