using System.Windows.Forms;
using MyOlap.Data;

namespace MyOlap.UI;

public class SettingsForm : Form
{
    private readonly CheckBox _chkOmitRows;
    private readonly CheckBox _chkOmitCols;
    private readonly CheckBox _chkPreserve;
    private readonly ComboBox _cbDisplay;
    private readonly Button _btnOk;
    private readonly Button _btnCancel;

    public ModelSettings Settings { get; private set; }

    public SettingsForm(ModelSettings current)
    {
        Settings = current;
        AutoScaleMode = AutoScaleMode.None;   // prevent DPI auto-scale from clipping text
        Text = "MyOlap – Settings";
        ClientSize = new System.Drawing.Size(820, 460);
        MinimumSize = new System.Drawing.Size(700, 400);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = false;
        MinimizeBox = false;

        var font = new System.Drawing.Font("Segoe UI", 10f);
        const int pad = 28;
        int y = 30;
        int chkW = ClientSize.Width - pad * 2;   // fill available width

        // ── Checkboxes ────────────────────────────────────────────────────────
        _chkOmitRows = new CheckBox
        {
            Text = "Omit Empty Rows",
            Left = pad, Top = y, Height = 36, Width = chkW,
            Font = font, AutoSize = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Checked = current.OmitEmptyRows
        };
        y += 50;

        _chkOmitCols = new CheckBox
        {
            Text = "Omit Empty Columns",
            Left = pad, Top = y, Height = 36, Width = chkW,
            Font = font, AutoSize = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Checked = current.OmitEmptyColumns
        };
        y += 50;

        _chkPreserve = new CheckBox
        {
            Text = "Preserve Excel Formulas and Text in Worksheet",
            Left = pad, Top = y, Height = 36, Width = chkW,
            Font = font, AutoSize = false,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
            Checked = current.PreserveFormulas
        };
        y += 66;

        // ── Member Display ────────────────────────────────────────────────────
        var lblDisplay = new Label
        {
            Text = "Member Display:",
            Left = pad, Top = y,
            AutoSize = true, Font = font
        };
        y += 34;

        _cbDisplay = new ComboBox
        {
            Left = pad, Top = y, Width = 360, Height = 32,
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = font
        };
        _cbDisplay.Items.AddRange(new object[] { "Name Only", "Description Only", "Name and Description" });
        _cbDisplay.SelectedIndex = current.MemberDisplay;

        // ── OK / Cancel ───────────────────────────────────────────────────────
        _btnOk = new Button
        {
            Text = "OK", Width = 110, Height = 36,
            DialogResult = DialogResult.OK,
            Font = new System.Drawing.Font("Segoe UI", 9.5f),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };
        _btnOk.Click += (_, _) =>
        {
            Settings = new ModelSettings
            {
                ModelId = current.ModelId,
                OmitEmptyRows = _chkOmitRows.Checked,
                OmitEmptyColumns = _chkOmitCols.Checked,
                MemberDisplay = _cbDisplay.SelectedIndex,
                PreserveFormulas = _chkPreserve.Checked
            };
        };

        _btnCancel = new Button
        {
            Text = "Cancel", Width = 110, Height = 36,
            DialogResult = DialogResult.Cancel,
            Font = new System.Drawing.Font("Segoe UI", 9.5f),
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };

        Controls.AddRange(new Control[]
        {
            _chkOmitRows, _chkOmitCols, _chkPreserve,
            lblDisplay, _cbDisplay,
            _btnOk, _btnCancel
        });
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        LayoutButtons();
        Resize += (_, _) => LayoutButtons();
    }

    private void LayoutButtons()
    {
        const int pad = 28;
        _btnCancel.Left = ClientSize.Width - pad - _btnCancel.Width;
        _btnCancel.Top = ClientSize.Height - pad - _btnCancel.Height;
        _btnOk.Left = _btnCancel.Left - _btnOk.Width - 12;
        _btnOk.Top = _btnCancel.Top;
    }
}
