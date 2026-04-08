using System.Windows.Forms;
using MyOlap.Core;
using MyOlap.Data;

namespace MyOlap.UI;

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

    public MemberPickerForm(long modelId, long initialDimensionId = 0)
    {
        AutoScaleMode = AutoScaleMode.Font;
        Text = "MyOlap \u2013 Pick Member";
        Width = 780;
        Height = 700;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        MinimumSize = new System.Drawing.Size(400, 400);

        int lx = 20;

        var lblDim = new Label
        {
            Text = "Dimension:", Left = lx, Top = 20, AutoSize = true
        };

        _cbDimension = new ComboBox
        {
            Left = lx, Top = 48, DropDownStyle = ComboBoxStyle.DropDownList,
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        };
        _cbDimension.SelectedIndexChanged += (_, _) => LoadTree();

        _tree = new TreeView
        {
            Left = lx, Top = 88, HideSelection = false, Scrollable = true,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        };

        _rbRow = new RadioButton
        {
            Text = "Place on Rows", Left = lx, AutoSize = true, Checked = true,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        };
        _rbCol = new RadioButton
        {
            Text = "Place on Columns", AutoSize = true,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        };

        _btnOk = new Button
        {
            Text = "OK", Width = 100, Height = 36,
            DialogResult = DialogResult.OK,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
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
            Text = "Cancel", Width = 110, Height = 36,
            DialogResult = DialogResult.Cancel,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        };

        Controls.AddRange(new Control[] { lblDim, _cbDimension, _tree, _rbRow, _rbCol, _btnOk, _btnCancel });
        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        LayoutControls();
        Resize += (_, _) => LayoutControls();

        _dimensions = SqliteRepository.Instance.GetDimensions(modelId);
        int initialIndex = 0;
        for (int i = 0; i < _dimensions.Count; i++)
        {
            _cbDimension.Items.Add(_dimensions[i].Name);
            if (_dimensions[i].Id == initialDimensionId)
                initialIndex = i;
        }
        if (_cbDimension.Items.Count > 0)
            _cbDimension.SelectedIndex = initialIndex;
    }

    private void LayoutControls()
    {
        int cw = ClientSize.Width - 40;
        int lx = 20;

        _cbDimension.Width = cw;
        _tree.Width = cw;
        _tree.Height = ClientSize.Height - 170;

        int radioY = ClientSize.Height - 72;
        _rbRow.Top = radioY;
        _rbRow.Left = lx;
        _rbCol.Top = radioY;
        _rbCol.Left = lx + 200;

        int btnY = ClientSize.Height - 56;
        _btnCancel.Left = lx + cw - _btnCancel.Width;
        _btnCancel.Top = btnY;
        _btnOk.Left = _btnCancel.Left - _btnOk.Width - 10;
        _btnOk.Top = btnY;
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