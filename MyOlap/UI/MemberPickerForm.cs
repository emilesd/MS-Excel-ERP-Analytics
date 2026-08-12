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
    private readonly Button _btnRemove;
    private readonly Button _btnMoveUp;
    private readonly Button _btnMoveDown;
    private readonly Button _btnClear;
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

    [Obsolete("Use SelectedMemberIds instead")]
    public long SelectedMemberId => _selectedMembers.Count > 0 ? _selectedMembers[0].Id : 0;

    private long _initialDimensionId;
    private List<(long Id, string Name)> _initialAxisMembers = new();
    private bool _axisPrePopulated;

    public MemberPickerForm(long modelId, long initialDimensionId = 0,
        List<(long Id, string Name)>? initialAxisMembers = null,
        bool initialPlaceOnRow = true)
    {
        AutoScaleMode = AutoScaleMode.Font;
        Text = "Pick Members";
        // Cap height to screen working area so OK/Cancel are always visible
        int screenH = Screen.PrimaryScreen?.WorkingArea.Height ?? 800;
        Width = 1380;
        Height = Math.Min(800, screenH - 40);
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
        _tree.NodeMouseDoubleClick += (_, e) => ToggleSingleNode(e.Node);
        _tree.AfterCheck += OnTreeAfterCheck;

        _btnRemove = new Button
        {
            Text = "Remove",
            Height = 36,
            Font = new System.Drawing.Font("Segoe UI", 9f)
        };
        _btnRemove.Click += (_, _) => RemoveSelectedFromList();

        _btnMoveUp = new Button { Text = "▲", Width = 44, Height = 34 };
        _btnMoveUp.Click += (_, _) => MoveSelected(-1);

        _btnMoveDown = new Button { Text = "▼", Width = 44, Height = 34 };
        _btnMoveDown.Click += (_, _) => MoveSelected(1);

        _btnClear = new Button { Text = "Clear", Width = 90, Height = 30 };
        _btnClear.Click += (_, _) => ClearSelected();

        var lblSelected = new Label { Text = "Selected Members:", AutoSize = true };
        _lbSelected = new ListBox { SelectionMode = SelectionMode.MultiExtended };

        _rbRow = new RadioButton { Text = "Place on Rows", AutoSize = true, Checked = initialPlaceOnRow };
        _rbCol = new RadioButton { Text = "Place on Columns", AutoSize = true, Checked = !initialPlaceOnRow };

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
            _tree, _btnRemove,
            lblSelected, _lbSelected,
            _btnMoveUp, _btnMoveDown, _btnClear,
            _rbRow, _rbCol, _btnOk, _btnCancel
        });

        AcceptButton = _btnOk;
        CancelButton = _btnCancel;

        LayoutControls();
        Resize += (_, _) => LayoutControls();

        _initialDimensionId = initialDimensionId;
        _initialAxisMembers = initialAxisMembers ?? new();

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

        int rightColW = 100;
        int gapW = 16;
        int leftPaneW = (cw - pad * 2 - gapW - rightColW - 16) / 2;
        int rightPaneX = pad + leftPaneW + gapW;
        int rightPaneW = cw - rightPaneX - rightColW - 16;

        int topY = pad;
        _cbDimension.Left = pad;
        _cbDimension.Top = topY;
        _cbDimension.Width = Math.Min(240, leftPaneW);

        int searchTop = topY + _cbDimension.Height + 12;
        _txtSearch.Left = pad;
        _txtSearch.Top = searchTop;
        _txtSearch.Width = leftPaneW;
        _txtSearch.PlaceholderText = "Enter a member name to search...";

        var lblSelected = Controls[4] as Label;
        lblSelected!.Left = rightPaneX;
        lblSelected.Top = searchTop;

        int treeTop = searchTop + 40;
        int bottomRowH = 56;
        int radioY = ch - bottomRowH;
        int paneH = radioY - 16 - treeTop;
        if (paneH < 60) paneH = 60;

        _tree.Left = pad;
        _tree.Top = treeTop;
        _tree.Width = leftPaneW;
        _tree.Height = paneH;

        _lbSelected.Left = rightPaneX;
        _lbSelected.Top = treeTop;
        _lbSelected.Width = rightPaneW;
        _lbSelected.Height = paneH;

        _btnRemove.Left = rightPaneX + rightPaneW - 160;
        _btnRemove.Top = searchTop - 2;
        _btnRemove.Width = 160;

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

        _rbRow.Left = pad;
        _rbRow.Top = radioY + 10;
        _rbCol.Left = _rbRow.Right + 24;
        _rbCol.Top = radioY + 10;

        _btnCancel.Left = cw - pad - _btnCancel.Width;
        _btnCancel.Top = radioY + 4;
        _btnOk.Left = _btnCancel.Left - _btnOk.Width - 12;
        _btnOk.Top = radioY + 4;
    }

    private void OnDimensionChanged()
    {
        _selectedMembers.Clear();
        if (!_axisPrePopulated && _cbDimension.SelectedIndex >= 0 &&
            _dimensions[_cbDimension.SelectedIndex].Id == _initialDimensionId &&
            _initialAxisMembers.Count > 0)
        {
            foreach (var m in _initialAxisMembers)
                _selectedMembers.Add(m);
            _axisPrePopulated = true;
        }
        RefreshSelectedList();
        _txtSearch.Text = "";
        LoadTree();
    }

    private HashSet<long> GetSelectedIds() => _selectedMembers.Select(m => m.Id).ToHashSet();

    private void LoadTree()
    {
        var selectedIds = GetSelectedIds();
        _suppressCheck = true;
        _tree.BeginUpdate();
        _tree.Nodes.Clear();
        if (_cbDimension.SelectedIndex < 0) { _tree.EndUpdate(); _suppressCheck = false; return; }

        var dim = _dimensions[_cbDimension.SelectedIndex];
        _currentRoots = DimensionTree.BuildTree(dim.Id);
        foreach (var root in _currentRoots)
            _tree.Nodes.Add(BuildNode(root, selectedIds));
        _tree.CollapseAll();
        _tree.EndUpdate();
        _suppressCheck = false;
    }

    private void ApplySearchFilter()
    {
        var query = _txtSearch.Text.Trim();
        var selectedIds = GetSelectedIds();
        _suppressCheck = true;
        _tree.BeginUpdate();
        _tree.Nodes.Clear();

        if (string.IsNullOrEmpty(query))
        {
            foreach (var root in _currentRoots)
                _tree.Nodes.Add(BuildNode(root, selectedIds));
            _tree.CollapseAll();
        }
        else
        {
            foreach (var root in _currentRoots)
            {
                var filtered = BuildFilteredNode(root, query, selectedIds);
                if (filtered != null)
                    _tree.Nodes.Add(filtered);
            }
            _tree.ExpandAll();
        }
        _tree.EndUpdate();
        _suppressCheck = false;
    }

    private static TreeNode? BuildFilteredNode(DimensionTreeNode dtNode, string query, HashSet<long> selectedIds)
    {
        bool selfMatch = dtNode.Member.DisplayName.Contains(query, StringComparison.OrdinalIgnoreCase)
            || (dtNode.Member.Description ?? "").Contains(query, StringComparison.OrdinalIgnoreCase);

        var childNodes = new List<TreeNode>();
        foreach (var child in dtNode.Children)
        {
            var filtered = BuildFilteredNode(child, query, selectedIds);
            if (filtered != null)
                childNodes.Add(filtered);
        }

        if (!selfMatch && childNodes.Count == 0)
            return null;

        var tn = new TreeNode(dtNode.Member.DisplayName)
        {
            Tag = dtNode.Member.Id,
            Checked = selectedIds.Contains(dtNode.Member.Id)
        };
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
        RebuildSelectionFromTree();
        RefreshSelectedList();
    }

    // Rebuilds _selectedMembers from all checked nodes in the tree.
    // Each node is independent — checking a parent does NOT cascade to children.
    private void RebuildSelectionFromTree()
    {
        var checkedNodes = new List<TreeNode>();
        CollectAllChecked(_tree.Nodes, checkedNodes);
        var checkedIds = checkedNodes.Select(n => (long)n.Tag!).ToHashSet();

        _selectedMembers.RemoveAll(m => !checkedIds.Contains(m.Id));

        var existingIds = _selectedMembers.Select(m => m.Id).ToHashSet();
        foreach (var node in checkedNodes)
        {
            if (node.Tag is long id && !existingIds.Contains(id))
                _selectedMembers.Add((id, node.Text));
        }
    }

    private static void CollectAllChecked(TreeNodeCollection nodes, List<TreeNode> results)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Checked && node.Tag is long)
                results.Add(node);
            CollectAllChecked(node.Nodes, results);
        }
    }

    private static void SetChildrenChecked(TreeNode node, bool isChecked)
    {
        foreach (TreeNode child in node.Nodes)
        {
            child.Checked = isChecked;
            SetChildrenChecked(child, isChecked);
        }
    }

    private void ToggleSingleNode(TreeNode? node)
    {
        if (node == null) return;
        _suppressCheck = true;
        node.Checked = !node.Checked;
        _suppressCheck = false;
        RebuildSelectionFromTree();
        RefreshSelectedList();
    }

    private void RemoveSelectedFromList()
    {
        var indices = _lbSelected.SelectedIndices.Cast<int>().OrderByDescending(i => i).ToList();
        var removedIds = indices.Select(i => _selectedMembers[i].Id).ToHashSet();
        _suppressCheck = true;
        UncheckNodes(_tree.Nodes, removedIds);
        _suppressCheck = false;
        foreach (var idx in indices)
            _selectedMembers.RemoveAt(idx);
        RefreshSelectedList();
    }

    private static void UncheckNodes(TreeNodeCollection nodes, HashSet<long>? ids)
    {
        foreach (TreeNode node in nodes)
        {
            if (ids == null || (node.Tag is long id && ids.Contains(id)))
                node.Checked = false;
            UncheckNodes(node.Nodes, ids);
        }
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
        _suppressCheck = true;
        UncheckNodes(_tree.Nodes, null);
        _suppressCheck = false;
        RefreshSelectedList();
    }

    private void RefreshSelectedList()
    {
        _lbSelected.Items.Clear();
        foreach (var (_, name) in _selectedMembers)
            _lbSelected.Items.Add(name);
        _btnOk.Enabled = _selectedMembers.Count > 0;
    }

    private static TreeNode BuildNode(DimensionTreeNode dtNode, HashSet<long> selectedIds)
    {
        var tn = new TreeNode(dtNode.Member.DisplayName)
        {
            Tag = dtNode.Member.Id,
            Checked = selectedIds.Contains(dtNode.Member.Id)
        };
        foreach (var child in dtNode.Children)
            tn.Nodes.Add(BuildNode(child, selectedIds));
        return tn;
    }
}
