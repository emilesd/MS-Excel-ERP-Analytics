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
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        Text = "Drill Down Options";
        Width = 400;
        Height = 240;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var rb1 = new RadioButton
        {
            Text = "Next Generation (children only)",
            Left = 20, Top = 18, Width = 340, Checked = true
        };
        var rb2 = new RadioButton
        {
            Text = "All Generations (full subtree)",
            Left = 20, Top = 50, Width = 340
        };
        var rb3 = new RadioButton
        {
            Text = "Base Generation Only (leaves)",
            Left = 20, Top = 82, Width = 340
        };

        var btnOk = new Button
        {
            Text = "OK", Left = 150, Top = 126, Width = 100, Height = 34,
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
            Text = "Cancel", Left = 260, Top = 126, Width = 100, Height = 34,
            DialogResult = DialogResult.Cancel
        };

        Controls.AddRange(new Control[] { rb1, rb2, rb3, btnOk, btnCancel });
        AcceptButton = btnOk;
        CancelButton = btnCancel;
    }
}
