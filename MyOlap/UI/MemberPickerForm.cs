using System.Windows.Forms;
using MyOlap.Core;
using MyOlap.Data;

namespace MyOlap.UI;

/// <summary>
/// Tree-view dialog for picking a dimension member.
/// Displays the full parent-child hierarchy.
/// </summary>
public class MemberPickerForm : Form
{
    private readonly TreeView _tree;
    private readonly ComboBox _cbDimension;
    private readonly RadioButton _rbRow;
    private readonly RadioButton _rbCol;
    private readonly Button _btnOk;
    private readonly Button _btnCancel;
    private List<Dimension> _dimensions = new();

    public long SelectedMemberId { get; private set; }
    public long SelectedDimensionId { get; private set; }
    public bool PlaceOnRow => _rbRow.Checked;

    public MemberPickerForm(long modelId)
    {
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        Text = "MyOlap \u2013 Pick Member";
        Width = 500;
        Height = 540;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblDim = new Label { Text = "Dimension:", Left = 12, Top = 14, Width = 90 };
        _cbDimension = new ComboBox
        {
            Left = 106, Top = 12, Width = 360, DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cbDimension.SelectedIndexChanged += (_, _) => LoadTree();

        _tree = new TreeView
        {
            Left = 12, Top = 46, Width = 456, Height = 340,
            HideSelection = false
        };

        var pnlPlacement = new Panel { Left = 12, Top = 396, Width = 456, Height = 34 };
        _rbRow = new RadioButton { Text = "Place on Rows", Left = 0, Top = 6, Width = 180, Checked = true };
        _rbCol = new RadioButton { Text = "Place on Columns", Left = 190, Top = 6, Width = 200 };
        pnlPlacement.Controls.Add(_rbRow);
        pnlPlacement.Controls.Add(_rbCol);

        _btnOk = new Button
        {
            Text = "OK", Left = 260, Top = 440, Width = 100, Height = 34,
            DialogResult = DialogResult.OK
        };
        _btnOk.Click += (_, _) =>
        {
            if (_tree.SelectedNode?.Tag is long id)
                SelectedMemberId = id;
            if (_cbDimension.SelectedIndex >= 0)
                SelectedDimensionId = _dimensions[_cbDimension.SelectedIndex].Id;
        };

        _btnCancel = new Button
        {
            Text = "Cancel", Left = 370, Top = 440, Width = 100, Height = 34,
            DialogResult = DialogResult.Cancel
        };

        Controls.AddRange(new Control[] { lblDim, _cbDimension, _tree, pnlPlacement, _btnOk, _btnCancel });
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        _dimensions = SqliteRepository.Instance.GetDimensions(modelId);
        foreach (var d in _dimensions)
            _cbDimension.Items.Add(d.Name);
        if (_cbDimension.Items.Count > 0)
            _cbDimension.SelectedIndex = 0;
    }

    private void LoadTree()
    {
        _tree.Nodes.Clear();
        if (_cbDimension.SelectedIndex < 0) return;

        var dim = _dimensions[_cbDimension.SelectedIndex];
        var roots = DimensionTree.BuildTree(dim.Id);
        foreach (var root in roots)
            _tree.Nodes.Add(BuildNode(root));
        _tree.ExpandAll();
    }

    private static TreeNode BuildNode(DimensionTreeNode dtNode)
    {
        var tn = new TreeNode(dtNode.Member.Name) { Tag = dtNode.Member.Id };
        foreach (var child in dtNode.Children)
            tn.Nodes.Add(BuildNode(child));
        return tn;
    }
}
