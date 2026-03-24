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
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        Text = "MyOlap \u2013 Settings";
        Width = 440;
        Height = 280;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        _chkOmitRows = new CheckBox
        {
            Text = "Omit Empty Rows",
            Left = 20, Top = 20, Width = 360,
            Checked = current.OmitEmptyRows
        };

        _chkOmitCols = new CheckBox
        {
            Text = "Omit Empty Columns",
            Left = 20, Top = 52, Width = 360,
            Checked = current.OmitEmptyColumns
        };

        var lblDisplay = new Label { Text = "Member Display:", Left = 20, Top = 92, Width = 130 };
        _cbDisplay = new ComboBox
        {
            Left = 155, Top = 90, Width = 240,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cbDisplay.Items.AddRange(new object[] { "Name Only", "Description Only", "Name and Description" });
        _cbDisplay.SelectedIndex = current.MemberDisplay;

        _btnOk = new Button
        {
            Text = "OK", Left = 190, Top = 150, Width = 100, Height = 34,
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
            Text = "Cancel", Left = 300, Top = 150, Width = 100, Height = 34,
            DialogResult = DialogResult.Cancel
        };

        Controls.AddRange(new Control[] { _chkOmitRows, _chkOmitCols, lblDisplay, _cbDisplay, _btnOk, _btnCancel });
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;
    }
}
