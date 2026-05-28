using System.Windows.Forms;
using MyOlap.Core;
using MyOlap.Data;

namespace MyOlap.UI;

public class MemberPickerForm : Form
{
    private readonly ComboBox _cbDimension;
    private readonly TextBox _txtSearch;
    private readonly TreeView _tree;
    private readonly ListBox _lbSelected;
    private readonly Button _btnAdd;
    private readonly Button _btnRemove;
    private readonly Button _btnMoveUp;
    private readonly Button _btnMoveDown;
    private readonly Button _btnClear;
    private readonly Button _btnRefresh;
    private readonly RadioButton _rbRow;
    private readonly RadioButton _rbCol;
    private readonly Button _btnOk;
    private readonly Button _btnCancel;

    private List<Dimension> _dimensions = new();
    private List<DimensionTreeNode> _currentRoots = new();
    private readonly List<(long Id, string Name)> _selectedMembers = new();

    public List<long> SelectedMemberIds => _selectedMembers.Select(m => m.Id).ToList();
    public long SelectedDimensionId { get; private set; }
    public bool PlaceOnRow => _rbRow.Checked;
    public bool RefreshRequested { get; private set; }

    [Obsolete("Use SelectedMemberIds instead")]
    public long SelectedMemberId => _selectedMembers.Count > 0 ? _selectedMembers[0].Id : 0;

    public MemberPickerForm(long modelId, long initialDimensionId = 0)
    {
        AutoScaleMode = AutoScaleMode.Font;
        Text = "Member Selection";
        Width = 1380;
        Height = 1400;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        MinimumSize = new System.Drawing.Size(780, 520);

        _cbDimension = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        _cbDimension.SelectedIndexChanged += (_, _) => OnDimensionChanged();

        _txtSearch = new TextBox();
        _txtSearch.TextChanged += (_, _) => ApplySearchFilter();

        _tree = new TreeView
        {
            CheckBoxes = true,
            HideSelection = false,
            Scrollable = true,
            ShowPlusMinus = true,
            ShowLines = true,
            ShowRootLines = true
        };
        _tree.NodeMouseDoubleClick += (_, e) => AddSingleNodeFromTree(e.Node);
        _tree.AfterCheck += OnTreeAfterCheck;

        _btnAdd = new Button { Text = "\u25B6", Width = 44, Height = 36, Font = new System.Drawing.Font("Segoe UI", 12f) };
        _btnAdd.Click += (_, _) => AddCheckedFromTree();

        _btnRemove = new Button { Text = "\u25C0", Width = 44, Height = 36, Font = new System.Drawing.Font("Segoe UI", 12f) };
        _btnRemove.Click += (_, _) => RemoveSelectedFromList();

        _btnMoveUp = new Button { Text = "\u25B2", Width = 44, Height = 34 };
        _btnMoveUp.Click += (_, _) => MoveSelected(-1);

        _btnMoveDown = new Button { Text = "\u25BC", Width = 44, Height = 34 };
        _btnMoveDown.Click += (_, _) => MoveSelected(1);

        _btnClear = new Button { Text = "Clear", Width = 90, Height = 30 };
        _btnClear.Click += (_, _) => ClearSelected();

        _btnRefresh = new Button { Text = "Refresh", Width = 90, Height = 30 };
        _btnRefresh.Click += (_, _) => LoadTree();

        var lblSelected = new Label { Text = "Selected Members:", AutoSize = true };
        _lbSelected = new ListBox { SelectionMode = SelectionMode.MultiExtended };

        _rbRow = new RadioButton { Text = "Place on Rows", AutoSize = true, Checked = true };
        _rbCol = new RadioButton { Text = "Place on Columns", AutoSize = true };

        _btnOk = new Button { Text = "OK", Width = 100, Height = 36, DialogResult = DialogResult.OK };
        _btnOk.Click += (_, _) =>
        {
            if (_cbDimension.SelectedIndex >= 0)
                SelectedDimensionId = _dimensions[_cbDimension.SelectedIndex].Id;
        };
        _btnOk.Enabled = false;

        _btnCancel = new Button { Text = "Cancel", Width = 100, Height = 36, DialogResult = DialogResult.Cancel };

        Controls.AddRange(new Control[]
        {
            _cbDimension, _txtSearch,
            _tree, _btnAdd, _btnRemove,
            lblSelected, _lbSelected,
            _btnMoveUp, _btnMoveDown, _btnClear, _btnRefresh,
            _rbRow, _rbCol, _btnOk, _btnCancel
        });

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
        int pad = 20;
        int cw = ClientSize.Width;
        int ch = ClientSize.Height;

        // Pane width calculations
        int centerBtnW = 56;
        int rightColW = 100;
        int leftPaneW = (cw - pad * 2 - centerBtnW - rightColW - 36) / 2;
        int rightPaneX = pad + leftPaneW + centerBtnW + 20;
        int rightPaneW = cw - rightPaneX - rightColW - 16;

        // Row 1: Dimension dropdown
        int topY = pad;
        _cbDimension.Left = pad;
        _cbDimension.Top = topY;
        _cbDimension.Width = Math.Min(240, leftPaneW);

        // Row 2: Search box (within left pane area, full width)
        int searchTop = topY + _cbDimension.Height + 12;
        _txtSearch.Left = pad;
        _txtSearch.Top = searchTop;
        _txtSearch.Width = leftPaneW;
        _txtSearch.PlaceholderText = "Enter a member name to search...";

        // "Selected Members:" label
        var lblSelected = Controls[5] as Label;
        lblSelected!.Left = rightPaneX;
        lblSelected.Top = searchTop;

        // Main pane area
        int treeTop = searchTop + 40;
        int bottomRowH = 52;
        int radioY = ch - bottomRowH;
        int paneBottom = radioY - 16;
        int paneH = paneBottom - treeTop;

        // Left tree
        _tree.Left = pad;
        _tree.Top = treeTop;
        _tree.Width = leftPaneW;
        _tree.Height = paneH;

        // Center ▶/◀ buttons
        int centerX = pad + leftPaneW + 6;
        int midY = treeTop + (paneH / 2) - 44;
        _btnAdd.Left = centerX;
        _btnAdd.Top = midY;
        _btnAdd.Width = centerBtnW - 12;
        _btnRemove.Left = centerX;
        _btnRemove.Top = midY + 46;
        _btnRemove.Width = centerBtnW - 12;

        // Right selected list
        _lbSelected.Left = rightPaneX;
        _lbSelected.Top = treeTop;
        _lbSelected.Width = rightPaneW;
        _lbSelected.Height = paneH;

        // Far-right buttons
        int orderX = rightPaneX + rightPaneW + 10;
        _btnMoveUp.Left = orderX;
        _btnMoveUp.Top = treeTop + 6;
        _btnMoveUp.Width = rightColW;
        _btnMoveDown.Left = orderX;
        _btnMoveDown.Top = treeTop + 46;
        _btnMoveDown.Width = rightColW;
        _btnClear.Left = orderX;
        _btnClear.Top = treeTop + 110;
        _btnClear.Width = rightColW;
        _btnRefresh.Left = orderX;
        _btnRefresh.Top = treeTop + 148;
        _btnRefresh.Width = rightColW;

        // Bottom row
        _rbRow.Left = pad;
        _rbRow.Top = radioY;
        _rbCol.Left = _rbRow.Right + 24;
        _rbCol.Top = radioY;

        _btnCancel.Left = cw - pad - _btnCancel.Width;
        _btnCancel.Top = radioY;
        _btnOk.Left = _btnCancel.Left - _btnOk.Width - 12;
        _btnOk.Top = radioY;
    }

    private void OnDimensionChanged()
    {
        _selectedMembers.Clear();
        RefreshSelectedList();
        _txtSearch.Text = "";
        LoadTree();
    }

    private void LoadTree()
    {
        _tree.BeginUpdate();
        _tree.Nodes.Clear();
        if (_cbDimension.SelectedIndex < 0) { _tree.EndUpdate(); return; }

        var dim = _dimensions[_cbDimension.SelectedIndex];
        _currentRoots = DimensionTree.BuildTree(dim.Id);
        foreach (var root in _currentRoots)
            _tree.Nodes.Add(BuildNode(root));
        _tree.ExpandAll();
        _tree.EndUpdate();
    }

    private void ApplySearchFilter()
    {
        var query = _txtSearch.Text.Trim();
        _tree.BeginUpdate();
        _tree.Nodes.Clear();

        if (string.IsNullOrEmpty(query))
        {
            foreach (var root in _currentRoots)
                _tree.Nodes.Add(BuildNode(root));
            _tree.ExpandAll();
        }
        else
        {
            foreach (var root in _currentRoots)
            {
                var filtered = BuildFilteredNode(root, query);
                if (filtered != null)
                    _tree.Nodes.Add(filtered);
            }
            _tree.ExpandAll();
        }
        _tree.EndUpdate();
    }

    private static TreeNode? BuildFilteredNode(DimensionTreeNode dtNode, string query)
    {
        bool selfMatch = dtNode.Member.Name.Contains(query, StringComparison.OrdinalIgnoreCase)
            || (dtNode.Member.Description ?? "").Contains(query, StringComparison.OrdinalIgnoreCase);

        var childNodes = new List<TreeNode>();
        foreach (var child in dtNode.Children)
        {
            var filtered = BuildFilteredNode(child, query);
            if (filtered != null)
                childNodes.Add(filtered);
        }

        if (!selfMatch && childNodes.Count == 0)
            return null;

        var tn = new TreeNode(dtNode.Member.Name) { Tag = dtNode.Member.Id };
        if (selfMatch)
            tn.BackColor = System.Drawing.Color.LightYellow;
        foreach (var cn in childNodes)
            tn.Nodes.Add(cn);
        return tn;
    }

    private bool _suppressCheck;

    private void OnTreeAfterCheck(object? sender, TreeViewEventArgs e)
    {
        if (_suppressCheck || e.Node == null) return;
        _suppressCheck = true;
        SetChildrenChecked(e.Node, e.Node.Checked);
        _suppressCheck = false;
    }

    private static void SetChildrenChecked(TreeNode node, bool isChecked)
    {
        foreach (TreeNode child in node.Nodes)
        {
            child.Checked = isChecked;
            SetChildrenChecked(child, isChecked);
        }
    }

    private void AddCheckedFromTree()
    {
        var checkedNodes = new List<TreeNode>();
        CollectCheckedLeaves(_tree.Nodes, checkedNodes);

        if (checkedNodes.Count == 0 && _tree.SelectedNode?.Tag is long id)
        {
            AddSingleNodeFromTree(_tree.SelectedNode);
            return;
        }

        foreach (var node in checkedNodes)
        {
            if (node.Tag is not long nid) continue;
            if (_selectedMembers.Any(m => m.Id == nid)) continue;
            _selectedMembers.Add((nid, node.Text));
        }
        RefreshSelectedList();
    }

    private static void CollectCheckedLeaves(TreeNodeCollection nodes, List<TreeNode> results)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Checked && node.Tag is long)
                results.Add(node);
            else
                CollectCheckedLeaves(node.Nodes, results);
        }
    }

    private void AddSingleNodeFromTree(TreeNode? node)
    {
        if (node?.Tag is not long id) return;
        if (_selectedMembers.Any(m => m.Id == id)) return;
        _selectedMembers.Add((id, node.Text));
        RefreshSelectedList();
    }

    private void RemoveSelectedFromList()
    {
        var indices = _lbSelected.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();
        foreach (var idx in indices)
            _selectedMembers.RemoveAt(idx);
        RefreshSelectedList();
    }

    private void MoveSelected(int direction)
    {
        if (_lbSelected.SelectedIndex < 0) return;
        int idx = _lbSelected.SelectedIndex;
        int newIdx = idx + direction;
        if (newIdx < 0 || newIdx >= _selectedMembers.Count) return;

        (_selectedMembers[idx], _selectedMembers[newIdx]) = (_selectedMembers[newIdx], _selectedMembers[idx]);
        RefreshSelectedList();
        _lbSelected.SelectedIndex = newIdx;
    }

    private void ClearSelected()
    {
        _selectedMembers.Clear();
        RefreshSelectedList();
    }

    private void RefreshSelectedList()
    {
        _lbSelected.Items.Clear();
        foreach (var (_, name) in _selectedMembers)
            _lbSelected.Items.Add(name);
        _btnOk.Enabled = _selectedMembers.Count > 0;
    }

    private static TreeNode BuildNode(DimensionTreeNode dtNode)
    {
        var tn = new TreeNode(dtNode.Member.Name) { Tag = dtNode.Member.Id };
        foreach (var child in dtNode.Children)
            tn.Nodes.Add(BuildNode(child));
        return tn;
    }
}
