using System.Windows.Forms;
using MyOlap.Core;
using MyOlap.Data;

namespace MyOlap.UI;

/// <summary>
/// Admin dialog for managing model dimensions and their members.
/// Covers: Manage Structure, Preview Tree, and Save Structure.
/// </summary>
public class ManageStructureForm : Form
{
    private readonly long _modelId;
    private readonly ListBox _lbDimensions;
    private readonly TreeView _tvMembers;
    private readonly Button _btnAddDim;
    private readonly Button _btnAddMember;
    private readonly Button _btnRemoveMember;
    private readonly Button _btnClose;
    private List<Dimension> _dims = new();

    public ManageStructureForm(long modelId)
    {
        _modelId = modelId;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoScaleDimensions = new SizeF(96F, 96F);
        Text = "MyOlap \u2013 Manage Model Structure";
        Width = 820;
        Height = 560;
        StartPosition = FormStartPosition.CenterScreen;

        var lblDim = new Label { Text = "Dimensions:", Left = 12, Top = 12, Width = 200 };
        _lbDimensions = new ListBox { Left = 12, Top = 34, Width = 220, Height = 400 };
        _lbDimensions.SelectedIndexChanged += (_, _) => LoadMembers();

        _btnAddDim = new Button { Text = "Add Dimension", Left = 12, Top = 444, Width = 150, Height = 34 };
        _btnAddDim.Click += BtnAddDim_Click;

        var lblMem = new Label { Text = "Members (Hierarchy):", Left = 248, Top = 12, Width = 220 };
        _tvMembers = new TreeView { Left = 248, Top = 34, Width = 540, Height = 400 };

        _btnAddMember = new Button { Text = "Add Member", Left = 248, Top = 444, Width = 140, Height = 34 };
        _btnAddMember.Click += BtnAddMember_Click;

        _btnRemoveMember = new Button { Text = "Remove Member", Left = 398, Top = 444, Width = 140, Height = 34 };
        _btnRemoveMember.Click += BtnRemoveMember_Click;

        _btnClose = new Button
        {
            Text = "Close", Left = 688, Top = 444, Width = 100, Height = 34,
            DialogResult = DialogResult.OK
        };

        Controls.AddRange(new Control[]
        {
            lblDim, _lbDimensions, _btnAddDim,
            lblMem, _tvMembers, _btnAddMember, _btnRemoveMember, _btnClose
        });

        LoadDimensions();
    }

    private void LoadDimensions()
    {
        _dims = SqliteRepository.Instance.GetDimensions(_modelId);
        _lbDimensions.Items.Clear();
        foreach (var d in _dims)
        {
            var typeTag = d.DimType != DimensionType.UserDefined ? $" [{d.DimType}]" : "";
            _lbDimensions.Items.Add($"{d.Name}{typeTag}");
        }
        if (_lbDimensions.Items.Count > 0)
            _lbDimensions.SelectedIndex = 0;
    }

    private void LoadMembers()
    {
        _tvMembers.Nodes.Clear();
        if (_lbDimensions.SelectedIndex < 0) return;
        var dim = _dims[_lbDimensions.SelectedIndex];
        var roots = DimensionTree.BuildTree(dim.Id);
        foreach (var root in roots)
            _tvMembers.Nodes.Add(BuildNode(root));
        _tvMembers.ExpandAll();
    }

    private static TreeNode BuildNode(DimensionTreeNode n)
    {
        var label = string.IsNullOrEmpty(n.Member.Description)
            ? n.Member.Name
            : $"{n.Member.Name} – {n.Member.Description}";
        var tn = new TreeNode(label) { Tag = n.Member.Id };
        foreach (var c in n.Children)
            tn.Nodes.Add(BuildNode(c));
        return tn;
    }

    private void BtnAddDim_Click(object? sender, EventArgs e)
    {
        var name = PromptInput("New Dimension", "Dimension name:");
        if (string.IsNullOrWhiteSpace(name)) return;

        var mgr = new ModelManager();
        var dim = mgr.AddDimension(_modelId, name);
        if (dim == null)
        {
            MessageBox.Show("Maximum of 12 dimensions reached.", "Limit",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        LoadDimensions();
    }

    private void BtnAddMember_Click(object? sender, EventArgs e)
    {
        if (_lbDimensions.SelectedIndex < 0) return;
        var dim = _dims[_lbDimensions.SelectedIndex];

        var name = PromptInput("New Member", "Member name:");
        if (string.IsNullOrWhiteSpace(name)) return;
        var desc = PromptInput("New Member", "Description (optional):");

        long? parentId = null;
        if (_tvMembers.SelectedNode?.Tag is long pid)
            parentId = pid;

        var mgr = new ModelManager();
        mgr.AddMember(dim.Id, name, desc ?? "", parentId);
        LoadMembers();
    }

    private void BtnRemoveMember_Click(object? sender, EventArgs e)
    {
        if (_tvMembers.SelectedNode?.Tag is not long memberId) return;
        var result = MessageBox.Show("Delete this member and its children?", "Confirm",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (result == DialogResult.Yes)
        {
            SqliteRepository.Instance.DeleteMember(memberId);
            LoadMembers();
        }
    }

    private static string? PromptInput(string title, string prompt)
    {
        var form = new Form
        {
            AutoScaleMode = AutoScaleMode.Dpi,
            AutoScaleDimensions = new SizeF(96F, 96F),
            Text = title, Width = 460, Height = 220,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false
        };
        var lbl = new Label { Text = prompt, Left = 16, Top = 16, Width = 400 };
        var txt = new TextBox { Left = 16, Top = 50, Width = 400 };
        var ok = new Button { Text = "OK", Left = 220, Top = 100, Width = 100, Height = 36, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", Left = 330, Top = 100, Width = 100, Height = 36, DialogResult = DialogResult.Cancel };
        form.Controls.AddRange(new Control[] { lbl, txt, ok, cancel });
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }
}
