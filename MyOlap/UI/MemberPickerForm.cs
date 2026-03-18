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
        Text = "MyOlap – Pick Member";
        Width = 420;
        Height = 500;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var lblDim = new Label { Text = "Dimension:", Left = 12, Top = 12, Width = 80 };
        _cbDimension = new ComboBox
        {
            Left = 96, Top = 10, Width = 300, DropDownStyle = ComboBoxStyle.DropDownList
        };
        _cbDimension.SelectedIndexChanged += (_, _) => LoadTree();

        _tree = new TreeView
        {
            Left = 12, Top = 44, Width = 380, Height = 320,
            HideSelection = false
        };

        var pnlPlacement = new Panel { Left = 12, Top = 374, Width = 380, Height = 30 };
        _rbRow = new RadioButton { Text = "Place on Rows", Left = 0, Top = 4, Width = 140, Checked = true };
        _rbCol = new RadioButton { Text = "Place on Columns", Left = 150, Top = 4, Width = 160 };
        pnlPlacement.Controls.Add(_rbRow);
        pnlPlacement.Controls.Add(_rbCol);

        _btnOk = new Button
        {
            Text = "OK", Left = 220, Top = 415, Width = 80, Height = 30,
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
            Text = "Cancel", Left = 310, Top = 415, Width = 80, Height = 30,
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
