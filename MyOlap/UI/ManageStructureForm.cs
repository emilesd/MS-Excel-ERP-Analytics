using System.Windows.Forms;
using MyOlap.Core;
using MyOlap.Data;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace MyOlap.UI;

public class ManageStructureForm : Form
{
    private readonly long _modelId;
    private readonly ListBox _lbDimensions;
    private readonly TreeView _tvMembers;
    private readonly Button _btnAddDim;
    private readonly Button _btnAddMember;
    private readonly Button _btnRemoveMember;
    private readonly Button _btnLoadDim;
    private readonly Button _btnExportPdf;
    private readonly Button _btnClose;
    private List<Dimension> _dims = new();

    public ManageStructureForm(long modelId)
    {
        _modelId = modelId;
        AutoScaleMode = AutoScaleMode.Font;
        Text = "MyOlap \u2013 Manage Model Structure";
        Width = 1300;
        Height = 700;
        MinimumSize = new System.Drawing.Size(1300, 700);
        StartPosition = FormStartPosition.CenterScreen;

        var lblDim = new Label { Text = "Dimensions:", Left = 12, Top = 12, AutoSize = true };
        _lbDimensions = new ListBox { Left = 12, Top = 42, Width = 220, Height = 500 };
        _lbDimensions.SelectedIndexChanged += (_, _) => LoadMembers();

        _btnAddDim = new Button { Text = "Add Dimension", Left = 12, Top = 554, Width = 220, Height = 36 };
        _btnAddDim.Click += BtnAddDim_Click;

        var lblMem = new Label { Text = "Members (Hierarchy):", Left = 248, Top = 12, AutoSize = true };
        _tvMembers = new TreeView { Left = 248, Top = 42, Width = 1020, Height = 500, Scrollable = true };

        var buttonPanel = new FlowLayoutPanel
        {
            Left = 248, Top = 554, Width = 1020, Height = 44,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };

        _btnAddMember = new Button
        {
            Text = "Add Member", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(80, 36),
            Margin = new Padding(0, 0, 6, 0)
        };
        _btnAddMember.Click += BtnAddMember_Click;

        _btnRemoveMember = new Button
        {
            Text = "Remove", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(70, 36),
            Margin = new Padding(0, 0, 6, 0)
        };
        _btnRemoveMember.Click += BtnRemoveMember_Click;

        _btnLoadDim = new Button
        {
            Text = "Load Dimension", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(80, 36),
            Margin = new Padding(0, 0, 6, 0)
        };
        _btnLoadDim.Click += BtnLoadDim_Click;

        _btnExportPdf = new Button
        {
            Text = "Export PDF", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(80, 36),
            Margin = new Padding(0, 0, 6, 0)
        };
        _btnExportPdf.Click += BtnExportPdf_Click;

        _btnClose = new Button
        {
            Text = "Close", Height = 36, AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowOnly,
            MinimumSize = new System.Drawing.Size(70, 36),
            DialogResult = DialogResult.OK
        };

        buttonPanel.Controls.AddRange(new Control[] { _btnAddMember, _btnRemoveMember, _btnLoadDim, _btnExportPdf, _btnClose });

        Controls.AddRange(new Control[]
        {
            lblDim, _lbDimensions, _btnAddDim,
            lblMem, _tvMembers, buttonPanel
        });

        LoadDimensions();
    }

    private void LoadDimensions()
    {
        _dims = SqliteRepository.Instance.GetDimensions(_modelId);
        _lbDimensions.Items.Clear();
        foreach (var d in _dims)
        {
            var typeTag = d.DimType != DimensionType.UserDefined ? " *" : "";
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
        var label = string.IsNullOrEmpty(n.Member.Description) || n.Member.Description == n.Member.Name
            ? n.Member.Name
            : $"{n.Member.Name} \u2013 {n.Member.Description}";
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

        long? parentId = null;
        if (_tvMembers.SelectedNode?.Tag is long pid)
            parentId = pid;

        var dlg = new Form
        {
            AutoScaleMode = AutoScaleMode.Font,
            
            Text = "Add Member",
            Width = 500, Height = 260,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false
        };

        var lblName = new Label { Text = "Name:", Left = 16, Top = 16, AutoSize = true };
        var txtName = new TextBox { Left = 130, Top = 14, Width = 340, MaxLength = 75 };

        var lblDesc = new Label { Text = "Description:", Left = 16, Top = 52, AutoSize = true };
        var txtDesc = new TextBox { Left = 130, Top = 50, Width = 340, MaxLength = 75 };

        var lblOp = new Label { Text = "Consol Operator:", Left = 16, Top = 88, AutoSize = true };
        var cbOp = new ComboBox
        {
            Left = 130, Top = 86, Width = 100, DropDownStyle = ComboBoxStyle.DropDownList
        };
        cbOp.Items.AddRange(new object[] { "+ (Add)", "- (Subtract)", "x (Ignore)" });
        cbOp.SelectedIndex = 0;

        var btnOk = new Button { Text = "OK", Left = 260, Top = 140, Width = 100, Height = 34, DialogResult = DialogResult.OK };
        var btnCancel = new Button { Text = "Cancel", Left = 370, Top = 140, Width = 100, Height = 34, DialogResult = DialogResult.Cancel };

        dlg.Controls.AddRange(new Control[] { lblName, txtName, lblDesc, txtDesc, lblOp, cbOp, btnOk, btnCancel });
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog() != DialogResult.OK || string.IsNullOrWhiteSpace(txtName.Text)) return;

        var consolOp = cbOp.SelectedIndex switch { 1 => "-", 2 => "x", _ => "+" };

        SqliteRepository.Instance.InsertMember(new Member
        {
            DimensionId = dim.Id,
            ParentId = parentId,
            Name = txtName.Text.Trim(),
            Description = txtDesc.Text.Trim(),
            Level = parentId.HasValue ? (SqliteRepository.Instance.GetMember(parentId.Value)?.Level ?? 0) + 1 : 0,
            SortOrder = SqliteRepository.Instance.GetMembers(dim.Id).Count,
            ConsolOperator = consolOp
        });
        LoadMembers();
    }

    private void BtnRemoveMember_Click(object? sender, EventArgs e)
    {
        if (_tvMembers.SelectedNode?.Tag is long memberId)
        {
            var result = MessageBox.Show("Delete this member and its children?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                var descendants = SqliteRepository.Instance.GetAllDescendants(memberId);
                foreach (var desc in descendants)
                    SqliteRepository.Instance.DeleteMember(desc.Id);
                SqliteRepository.Instance.DeleteMember(memberId);
                LoadMembers();
            }
            return;
        }

        if (_lbDimensions.SelectedIndex < 0) return;
        var dim = _dims[_lbDimensions.SelectedIndex];

        if (dim.DimType != DimensionType.UserDefined)
        {
            MessageBox.Show("Cannot delete predefined dimension. Predefined dimensions (View, Version, Year, Measure, Time) cannot be removed.",
                "Not Allowed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var confirmResult = MessageBox.Show("Delete dimension and all its members?", "Confirm",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (confirmResult == DialogResult.Yes)
        {
            SqliteRepository.Instance.ClearDimensionMembers(dim.Id);
            SqliteRepository.Instance.DeleteDimension(dim.Id);
            LoadDimensions();
        }
    }

    private void BtnLoadDim_Click(object? sender, EventArgs e)
    {
        using var form = new LoadDimensionForm(_modelId);
        if (form.ShowDialog(this) == DialogResult.OK)
        {
            LoadDimensions();
            LoadMembers();
        }
    }

    private void BtnExportPdf_Click(object? sender, EventArgs e)
    {
        if (_lbDimensions.SelectedIndex < 0)
        {
            MessageBox.Show("Select a dimension first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var dim = _dims[_lbDimensions.SelectedIndex];
        var roots = DimensionTree.BuildTree(dim.Id);
        if (roots.Count == 0)
        {
            MessageBox.Show("No members in this dimension.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        using var dlg = new SaveFileDialog
        {
            Filter = "PDF Files|*.pdf",
            FileName = $"{dim.Name}_Structure.pdf",
            Title = "Export Dimension Tree to PDF"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            ExportDimensionTreePdf(dim, roots, dlg.FileName);
            MessageBox.Show($"PDF exported to:\n{dlg.FileName}", "Done",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error exporting PDF: {ex.Message}", "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static void ExportDimensionTreePdf(Dimension dim, List<DimensionTreeNode> roots, string outputPath)
    {
        var doc = new PdfDocument();
        doc.Info.Title = $"{dim.Name} - Dimension Structure";

        var titleFont = new XFont("Arial", 14, XFontStyle.Bold);
        var parentFont = new XFont("Arial", 10, XFontStyle.Bold);
        var leafFont = new XFont("Arial", 10, XFontStyle.Regular);
        var smallFont = new XFont("Arial", 8, XFontStyle.Regular);

        double marginLeft = 40;
        double marginTop = 40;
        double lineHeight = 16;
        double indentWidth = 20;

        var page = doc.AddPage();
        page.Orientation = PdfSharpCore.PageOrientation.Portrait;
        var gfx = XGraphics.FromPdfPage(page);
        double y = marginTop;
        double pageBottom = page.Height - marginTop;

        gfx.DrawString($"Dimension: {dim.Name}", titleFont, XBrushes.Black, new XPoint(marginLeft, y));
        y += 22;
        gfx.DrawString($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm}", smallFont, XBrushes.DarkGray, new XPoint(marginLeft, y));
        y += 20;

        void RenderNode(DimensionTreeNode node, int depth)
        {
            if (y + lineHeight > pageBottom)
            {
                page = doc.AddPage();
                page.Orientation = PdfSharpCore.PageOrientation.Portrait;
                gfx = XGraphics.FromPdfPage(page);
                y = marginTop;
            }

            double x = marginLeft + depth * indentWidth;
            bool hasChildren = node.Children.Count > 0;
            var font = hasChildren ? parentFont : leafFont;
            var prefix = hasChildren ? "\u25BC " : "\u2022 ";
            var label = string.IsNullOrEmpty(node.Member.Description) || node.Member.Description == node.Member.Name
                ? node.Member.Name
                : $"{node.Member.Name} \u2013 {node.Member.Description}";

            gfx.DrawString($"{prefix}{label}", font, XBrushes.Black, new XPoint(x, y));
            y += lineHeight;

            foreach (var child in node.Children)
                RenderNode(child, depth + 1);
        }

        foreach (var root in roots)
            RenderNode(root, 0);

        doc.Save(outputPath);
    }

    private static string? PromptInput(string title, string prompt)
    {
        var form = new Form
        {
            AutoScaleMode = AutoScaleMode.Font,
            
            Text = title, Width = 460, Height = 220,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false
        };
        var lbl = new Label { Text = prompt, Left = 16, Top = 16, AutoSize = true };
        var txt = new TextBox { Left = 16, Top = 52, Width = 400, Height = 28 };
        var ok = new Button { Text = "OK", Left = 220, Top = 104, Width = 100, Height = 36, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", Left = 330, Top = 104, Width = 100, Height = 36, DialogResult = DialogResult.Cancel };
        form.Controls.AddRange(new Control[] { lbl, txt, ok, cancel });
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        return form.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }
}