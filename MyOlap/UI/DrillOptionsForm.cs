using System.Windows.Forms;
using MyOlap.Core;

namespace MyOlap.UI;

/// <summary>
/// Small dialog to choose the drill-down mode:
/// Next Generation, All Generations, or Base Generation Only.
/// </summary>
public class DrillOptionsForm : Form
{
    public DrillMode SelectedMode { get; private set; } = DrillMode.NextGeneration;

    public DrillOptionsForm()
    {
        Text = "Drill Down Options";
        Width = 320;
        Height = 220;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        MinimizeBox = false;

        var rb1 = new RadioButton
        {
            Text = "Next Generation (children only)",
            Left = 20, Top = 15, Width = 260, Checked = true
        };
        var rb2 = new RadioButton
        {
            Text = "All Generations (full subtree)",
            Left = 20, Top = 45, Width = 260
        };
        var rb3 = new RadioButton
        {
            Text = "Base Generation Only (leaves)",
            Left = 20, Top = 75, Width = 260
        };

        var btnOk = new Button
        {
            Text = "OK", Left = 100, Top = 115, Width = 80, Height = 28,
            DialogResult = DialogResult.OK
        };
        btnOk.Click += (_, _) =>
        {
            if (rb2.Checked) SelectedMode = DrillMode.AllGenerations;
            else if (rb3.Checked) SelectedMode = DrillMode.BaseOnly;
            else SelectedMode = DrillMode.NextGeneration;
        };

        var btnCancel = new Button
        {
            Text = "Cancel", Left = 190, Top = 115, Width = 80, Height = 28,
            DialogResult = DialogResult.Cancel
        };

        Controls.AddRange(new Control[] { rb1, rb2, rb3, btnOk, btnCancel });
        AcceptButton = btnOk;
        CancelButton = btnCancel;
    }
}
