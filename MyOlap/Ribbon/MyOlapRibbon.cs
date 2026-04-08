using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using MyOlap.Core;
using MyOlap.Data;
using MyOlap.Reports;
using MyOlap.UI;

namespace MyOlap.Ribbon;

/// <summary>
/// Excel-DNA ribbon controller implementing all MyOlap menu buttons.
/// Each callback maps directly to a Product Brief requirement.
/// </summary>
[ComVisible(true)]
public class MyOlapRibbon : ExcelRibbon
{
    private readonly OlapEngine _engine = OlapEngine.Instance;
    private IRibbonUI? _ribbonUi;

    public void OnRibbonLoad(IRibbonUI ribbonUI)
    {
        _ribbonUi = ribbonUI;
    }

    private void RefreshInfoLabels()
    {
        try { _ribbonUi?.InvalidateControl("lblActiveModel"); } catch { }
    }

    public override string GetCustomUI(string ribbonId)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='tabMyOlap' label='MyOlap'>
        <group id='grpModel' label='Model'>
          <button id='btnSelectModel'  label='Select Model'  size='large' imageMso='_3DStyle'  onAction='OnSelectModel'/>
          <button id='btnRefreshData'  label='Refresh Data'  size='large' imageMso='RefreshAll'              onAction='OnRefreshData'/>
        </group>
        <group id='grpNavigate' label='Navigate'>
          <button id='btnPickMember'   label='Pick Member'   size='large' imageMso='OrganizationChartInsert'         onAction='OnPickMember'/>
          <button id='btnDrillDown'    label='Drill Down'    size='normal' imageMso='OutlineShowDetail'              onAction='OnDrillDown'/>
          <button id='btnDrillUp'      label='Drill Up'      size='normal' imageMso='OutlineHideDetail'             onAction='OnDrillUp'/>
        </group>
        <group id='grpView' label='View'>
          <button id='btnSwapRowCol'   label='Swap to Row/Col' size='normal' imageMso='PivotTableReport'   onAction='OnSwapRowCol'/>
          <button id='btnKeepSelected' label='Keep Selected' size='normal' imageMso='FilterBySelection'   onAction='OnKeepSelected'/>
          <button id='btnRemoveSelected' label='Remove Selected' size='normal' imageMso='Delete'          onAction='OnRemoveSelected'/>
          <button id='btnUndoLast'     label='Undo Last'     size='normal' imageMso='Undo'                onAction='OnUndoLast'/>
        </group>
        <group id='grpAdmin' label='Admin'>
          <button id='btnManageModel'  label='Manage Model'  size='large' imageMso='DesignMode'        onAction='OnManageModel'/>
          <button id='btnClearData'    label='Clear Data'    size='normal' imageMso='RecordsDeleteRecord'  onAction='OnClearData'/>
          <button id='btnLoadData'     label='Load Data'     size='normal' imageMso='ImportTextFile'       onAction='OnLoadData'/>
          <button id='btnSettings'     label='Settings'      size='normal' imageMso='ControlProperties'    onAction='OnSettings'/>
        </group>

        <group id='grpInfo' label='Info'>
          <labelControl id='lblActiveModel' getLabel='GetActiveModelLabel'/>
          <labelControl id='lblVersion'     label='Version: v1.0'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public string GetActiveModelLabel(IRibbonControl control)
    {
        var name = _engine.ActiveModel?.Name;
        return string.IsNullOrEmpty(name) ? "Model: (none)" : $"Model: {name}";
    }

    #region Ribbon Callbacks

    public void OnSelectModel(IRibbonControl control)
    {
        try
        {
            using var form = new ModelBrowserForm();
            var owner = new Win32Window(GetExcelHwnd());
            if (form.ShowDialog(owner) != DialogResult.OK) return;

            if (form.CreateNew)
            {
                var name = PromptInput("New Model", "Model name:");
                if (string.IsNullOrWhiteSpace(name)) return;

                var mgr = new ModelManager();
                long modelId;
                if (form.CloneFromId > 0)
                    modelId = mgr.CloneModel(form.CloneFromId, name);
                else
                    modelId = mgr.CreateEmptyModel(name);
                _engine.SelectModel(modelId);
                WriteGridToSheet();
                RefreshInfoLabels();
            }
            else if (form.SelectedModelId > 0)
            {
                _engine.SelectModel(form.SelectedModelId);
                WriteGridToSheet();
                RefreshInfoLabels();
            }
        }
        catch (Exception ex)
        {
            var msg = $"Select Model error:\n{ex.Message}";
            var inner = ex.InnerException;
            while (inner != null)
            {
                msg += $"\n\nInner: {inner.Message}";
                inner = inner.InnerException;
            }
            msg += $"\n\nStack: {ex.StackTrace}";
            MessageBox.Show(msg, "MyOlap Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnRefreshData(IRibbonControl control)
    {
        if (_engine.ActiveModel == null)
        {
            ShowMessage("No model selected. Use 'Select Model' first.");
            return;
        }
        var sheetName = GetActiveSheetName();
        if (_engine.RestoreViewForSheet(sheetName))
        {
            WriteGridToSheet();
        }
        else
        {
            _engine.SelectModel(_engine.ActiveModel.Id);
            WriteGridToSheet();
        }
    }

    private string GetActiveSheetName()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            return app.ActiveSheet.Name ?? "Sheet1";
        }
        catch { return "Sheet1"; }
    }

    private bool IsActiveSheetBlank()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic ws = app.ActiveSheet;
            object val = ws.Cells[1, 1].Value;
            return val == null || string.IsNullOrEmpty(val.ToString());
        }
        catch { return true; }
    }

    public void OnPickMember(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        try
        {
            var (contextDimId, _) = GetSelectedCellMember();
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new MemberPickerForm(_engine.ActiveModel.Id, contextDimId);
            if (form.ShowDialog(owner) != DialogResult.OK) return;
            if (form.SelectedMemberId <= 0) return;

            _engine.PickMember(form.SelectedDimensionId, form.SelectedMemberId, form.PlaceOnRow);
            WriteGridToSheet();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Pick Member error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnDrillDown(IRibbonControl control)
    {
        if (_engine.ActiveModel == null || _engine.CurrentView == null)
        { ShowMessage("No model selected."); return; }

        var selected = GetAllSelectedCellMembers();
        if (selected.Count == 0)
        {
            var (d, m) = GetSelectedCellMember();
            if (d > 0 && m > 0) selected.Add((d, m));
        }
        if (selected.Count == 0)
        { ShowMessage("Select member cell(s) to drill down on."); return; }

        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new DrillOptionsForm();
            if (form.ShowDialog(owner) != DialogResult.OK) return;

            foreach (var (dimId, memberId) in selected.Distinct())
                _engine.DrillDown(dimId, memberId, form.SelectedMode);
            WriteGridToSheet();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Drill Down error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnDrillUp(IRibbonControl control)
    {
        if (_engine.ActiveModel == null || _engine.CurrentView == null)
        { ShowMessage("No model selected."); return; }

        var selected = GetAllSelectedCellMembers();
        if (selected.Count == 0)
        {
            var (d, m) = GetSelectedCellMember();
            if (d > 0 && m > 0) selected.Add((d, m));
        }
        if (selected.Count == 0)
        { ShowMessage("Select member cell(s) to drill up on."); return; }

        foreach (var (dimId, memberId) in selected.Distinct())
            _engine.DrillUp(dimId, memberId);
        WriteGridToSheet();
    }

    public void OnSwapRowCol(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        var (dimId, _) = GetSelectedCellMember();
        if (dimId > 0)
            _engine.SwapDimension(dimId);
        else
            _engine.SwapRowCol();
        WriteGridToSheet();
    }

    public void OnKeepSelected(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0 || memberId == 0) { ShowMessage("Select a member cell first."); return; }
        _engine.KeepSelected(dimId, memberId);
        WriteGridToSheet();
    }

    public void OnRemoveSelected(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0 || memberId == 0) { ShowMessage("Select a member cell first."); return; }
        _engine.RemoveSelected(dimId, memberId);
        WriteGridToSheet();
    }

    public void OnUndoLast(IRibbonControl control)
    {
        if (!_engine.CanUndo) { ShowMessage("Nothing to undo."); return; }
        _engine.Undo();
        WriteGridToSheet();
    }

    public void OnManageModel(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new ManageStructureForm(_engine.ActiveModel.Id);
            if (form.ShowDialog(owner) == DialogResult.OK)
                WriteGridToSheet();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Manage Model error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnLoadData(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new DataLoadForm(_engine.ActiveModel.Id);
            if (form.ShowDialog(owner) == DialogResult.OK)
                WriteGridToSheet();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Load Data error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnClearData(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            var form = new Form
            {
                AutoScaleMode = AutoScaleMode.Font,
                Text = "MyOlap \u2013 Clear Data",
                Width = 520, Height = 340,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MaximizeBox = false, MinimizeBox = false
            };

            var rbAll = new RadioButton { Text = "Clear All Data", Left = 20, Top = 18, Width = 420, Checked = true };
            var rbFilter = new RadioButton { Text = "Clear by View + Version + Year:", Left = 20, Top = 50, Width = 420 };

            var dims = SqliteRepository.Instance.GetDimensions(_engine.ActiveModel.Id);
            var viewDim = dims.FirstOrDefault(d => d.DimType == DimensionType.View);
            var versionDim = dims.FirstOrDefault(d => d.DimType == DimensionType.Version);
            var yearDim = dims.FirstOrDefault(d => d.DimType == DimensionType.Year);

            var viewMembers = viewDim != null ? SqliteRepository.Instance.GetMembers(viewDim.Id) : new List<Member>();
            var versionMembers = versionDim != null ? SqliteRepository.Instance.GetMembers(versionDim.Id) : new List<Member>();
            var yearMembers = yearDim != null ? SqliteRepository.Instance.GetMembers(yearDim.Id) : new List<Member>();

            var lblView = new Label { Text = "View:", Left = 40, Top = 80, AutoSize = true };
            var cbView = new ComboBox { Left = 40, Top = 110, Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
            var lblVersion = new Label { Text = "Version:", Left = 190, Top = 80, AutoSize = true };
            var cbVersion = new ComboBox { Left = 190, Top = 110, Width = 130, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };
            var lblYear = new Label { Text = "Year:", Left = 340, Top = 80, AutoSize = true };
            var cbYear = new ComboBox { Left = 340, Top = 110, Width = 140, DropDownStyle = ComboBoxStyle.DropDownList, Enabled = false };

            foreach (var m in viewMembers) cbView.Items.Add(m.Name);
            foreach (var m in versionMembers) cbVersion.Items.Add(m.Name);
            foreach (var m in yearMembers) cbYear.Items.Add(m.Name);

            if (cbView.Items.Count > 0) cbView.SelectedIndex = 0;
            if (cbVersion.Items.Count > 0) cbVersion.SelectedIndex = 0;
            if (cbYear.Items.Count > 0) cbYear.SelectedIndex = 0;

            rbFilter.CheckedChanged += (_, _) =>
            {
                cbView.Enabled = rbFilter.Checked;
                cbVersion.Enabled = rbFilter.Checked;
                cbYear.Enabled = rbFilter.Checked;
            };

            var btnOk = new Button { Text = "Clear", Left = 260, Top = 200, Width = 100, Height = 34, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Cancel", Left = 370, Top = 200, Width = 100, Height = 34, DialogResult = DialogResult.Cancel };

            form.Controls.AddRange(new Control[] { rbAll, rbFilter, lblView, cbView, lblVersion, cbVersion, lblYear, cbYear, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            if (form.ShowDialog(owner) != DialogResult.OK) return;

            var confirm = MessageBox.Show("Are you sure you want to clear data? This cannot be undone.",
                "Confirm Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            if (rbAll.Checked)
            {
                SqliteRepository.Instance.ClearFacts(_engine.ActiveModel.Id);
                ShowMessage("All data cleared.");
            }
            else
            {
                var dimOrder = dims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();
                long? selViewId = (cbView.SelectedIndex >= 0 && viewMembers.Count > cbView.SelectedIndex) ? viewMembers[cbView.SelectedIndex].Id : null;
                long? selVersionId = (cbVersion.SelectedIndex >= 0 && versionMembers.Count > cbVersion.SelectedIndex) ? versionMembers[cbVersion.SelectedIndex].Id : null;
                long? selYearId = (cbYear.SelectedIndex >= 0 && yearMembers.Count > cbYear.SelectedIndex) ? yearMembers[cbYear.SelectedIndex].Id : null;

                SqliteRepository.Instance.ClearFactsByFilter(
                    _engine.ActiveModel.Id, dimOrder,
                    selViewId, selVersionId, selYearId,
                    viewDim?.Id, versionDim?.Id, yearDim?.Id);

                ShowMessage("Filtered data cleared.");
            }

            WriteGridToSheet();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Clear Data error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnSettings(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            var current = SqliteRepository.Instance.GetSettings(_engine.ActiveModel.Id);
            using var form = new SettingsForm(current);
            if (form.ShowDialog(owner) == DialogResult.OK)
            {
                SqliteRepository.Instance.SaveSettings(form.Settings);
                WriteGridToSheet();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Settings error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnExportPdf(IRibbonControl control)
    {
        if (_engine.ActiveModel == null || _engine.CurrentView == null)
        { ShowMessage("No model/view active."); return; }

        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var dlg = new SaveFileDialog
            {
                Filter = "PDF Files|*.pdf",
                Title = "Export Report to PDF",
                FileName = $"MyOlap_{_engine.ActiveModel.Name}_{DateTime.Now:yyyyMMdd}.pdf"
            };
            if (dlg.ShowDialog(owner) != DialogResult.OK) return;

            var grid = _engine.BuildGrid();
            var builder = new ReportBuilder();
            var report = builder.BuildFromGrid(grid, _engine.ActiveModel.Name);
            var exporter = new PdfExporter();
            exporter.Export(report, dlg.FileName);
            ShowMessage($"Report exported to {dlg.FileName}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    #endregion

    #region Grid Rendering

    private static readonly BindingFlags GP = BindingFlags.GetProperty;
    private static readonly BindingFlags SP = BindingFlags.SetProperty;
    private static readonly BindingFlags IM = BindingFlags.InvokeMethod;

    /// <summary>
    /// Gets the active worksheet via COM reflection. Returns null if unavailable.
    /// </summary>
    private static object? GetActiveSheet()
    {
        object? xlApp = ExcelDnaUtil.Application;
        if (xlApp == null) return null;
        return xlApp.GetType().InvokeMember("ActiveSheet", GP, null, xlApp, null);
    }

    /// <summary>
    /// Sets a cell value using COM reflection (1-based row/col).
    /// </summary>
    private static void ComSetCell(object ws, int row1, int col1, object value)
    {
        var wsType = ws.GetType();
        object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row1, col1 });
        if (cell == null) return;
        cell.GetType().InvokeMember("Value2", SP, null, cell, new object[] { value });
    }

    /// <summary>
    /// Adds a comment/note to a cell for storing dimension/member metadata.
    /// </summary>
    private static void ComSetComment(object ws, int row1, int col1, string text)
    {
        try
        {
            var wsType = ws.GetType();
            object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row1, col1 });
            if (cell == null) return;
            var cellType = cell.GetType();
            try { cellType.InvokeMember("ClearComments", IM, null, cell, null); } catch { }
            object? comment = cellType.InvokeMember("AddComment", IM, null, cell, new object[] { text });
            if (comment != null)
                comment.GetType().InvokeMember("Visible", SP, null, comment, new object[] { false });
        }
        catch { }
    }

    private static void ComSetBold(object ws, int row1, int col1, bool bold)
    {
        try
        {
            var wsType = ws.GetType();
            object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row1, col1 });
            if (cell == null) return;
            object? font = cell.GetType().InvokeMember("Font", GP, null, cell, null);
            if (font != null)
                font.GetType().InvokeMember("Bold", SP, null, font, new object[] { bold });
        }
        catch { }
    }

    private static void ComSetNumberFormat(object ws, int row1, int col1, string format)
    {
        try
        {
            object? cell = ws.GetType().InvokeMember("Cells", GP, null, ws, new object[] { row1, col1 });
            if (cell != null)
                cell.GetType().InvokeMember("NumberFormat", SP, null, cell, new object[] { format });
        }
        catch { }
    }

    /// <summary>
    /// Clears all cells on the active worksheet using COM reflection.
    /// </summary>
    private static void ComClearSheet(object ws)
    {
        try
        {
            var wsType = ws.GetType();
            object? cells = wsType.InvokeMember("Cells", GP, null, ws, null);
            if (cells != null)
                cells.GetType().InvokeMember("Clear", IM, null, cells, null);
        }
        catch { }
    }

    /// <summary>
    /// Writes the OLAP grid to the active worksheet using COM via reflection.
    /// No dynamic keyword, no C API - purely reflection-based COM calls.
    /// </summary>
    private void WriteGridToSheet()
    {
        try { _engine.SaveViewForSheet(GetActiveSheetName()); } catch { }
        try
        {
            var grid = _engine.BuildGrid();
            var view = _engine.CurrentView;
            var modelName = _engine.ActiveModel?.Name ?? "Model";

            object? ws = GetActiveSheet();
            if (ws == null)
            {
                ShowMessage("No active worksheet found. Please open or create a worksheet first.");
                return;
            }

            ComClearSheet(ws);

            try
            {
                object? xlApp = ExcelDnaUtil.Application;
                if (xlApp != null)
                    xlApp.GetType().InvokeMember("DisplayCommentIndicator", SP, null, xlApp, new object[] { 0 });
            }
            catch { }

            if (grid.RowHeaders.Count == 0 && grid.ColHeaders.Count == 0)
            {
                ComSetCell(ws, 1, 1, $"Model '{modelName}' is ready.");
                ComSetCell(ws, 2, 1, "Use Manage Model to add dimensions/members, then Refresh Data.");
                return;
            }

            int povRowCount = 0;
            if (view != null && view.PovSelections.Count > 0)
            {
                var allDims = SqliteRepository.Instance.GetDimensions(view.ModelId);
                int povCol = 1;
                foreach (var kvp in view.PovSelections)
                {
                    var dim = allDims.FirstOrDefault(d => d.Id == kvp.Key);
                    var member = SqliteRepository.Instance.GetMember(kvp.Value);
                    if (dim == null || member == null) continue;
                    var label = $"{dim.Name}: {grid.FormatMember(member)}";
                    ComSetCell(ws, 1, povCol, label);
                    ComSetBold(ws, 1, povCol, true);
                    ComSetComment(ws, 1, povCol, $"DIM:{dim.Id}|MBR:{member.Id}");
                    povCol++;
                }
                povRowCount = 2;
            }

            int headerRows = (grid.ColDimensionNames.Count > 0
                ? grid.ColDimensionNames.Count + 1 : 1) + povRowCount;
            int headerCols = grid.RowDimensionNames.Count > 0
                ? grid.RowDimensionNames.Count : 1;

            for (int i = 0; i < grid.RowDimensionNames.Count; i++)
            {
                ComSetCell(ws, 1 + povRowCount, i + 1, grid.RowDimensionNames[i]);
                ComSetBold(ws, 1 + povRowCount, i + 1, true);
            }

            for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
            {
                var combo = grid.ColHeaders[cIdx];
                for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                {
                    ComSetCell(ws, dIdx + 1 + povRowCount, headerCols + cIdx + 1, grid.FormatMember(combo[dIdx]));
                    ComSetBold(ws, dIdx + 1 + povRowCount, headerCols + cIdx + 1, true);
                    if (view != null && dIdx < view.ColAxes.Count)
                        ComSetComment(ws, dIdx + 1 + povRowCount, headerCols + cIdx + 1,
                            $"DIM:{view.ColAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}");
                }
            }

            for (int rIdx = 0; rIdx < grid.RowHeaders.Count; rIdx++)
            {
                var combo = grid.RowHeaders[rIdx];
                for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                {
                    ComSetCell(ws, headerRows + rIdx + 1, dIdx + 1, grid.FormatMember(combo[dIdx]));
                    if (SqliteRepository.Instance.GetChildren(combo[dIdx].Id).Count > 0)
                        ComSetBold(ws, headerRows + rIdx + 1, dIdx + 1, true);
                    if (view != null && dIdx < view.RowAxes.Count)
                        ComSetComment(ws, headerRows + rIdx + 1, dIdx + 1,
                            $"DIM:{view.RowAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}");
                }
            }

            for (int rIdx = 0; rIdx < grid.RowHeaders.Count; rIdx++)
            {
                for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
                {
                    var val = grid.Values[rIdx, cIdx];
                    if (val.HasValue)
                    {
                        int dataRow = headerRows + rIdx + 1;
                        int dataCol = headerCols + cIdx + 1;
                        ComSetCell(ws, dataRow, dataCol, (double)val.Value);
                        ComSetNumberFormat(ws, dataRow, dataCol, "#,##0.00");
                    }
                }
            }

            ComAutoFit(ws);
        }
        catch (Exception ex)
        {
            ShowMessage($"Error rendering grid: {ex.Message}\n\nStack: {ex.StackTrace}");
        }
    }

    /// <summary>
    /// Auto-fits all columns on the worksheet.
    /// </summary>
    private static void ComAutoFit(object ws)
    {
        try
        {
            var wsType = ws.GetType();
            object? columns = wsType.InvokeMember("Columns", GP, null, ws, null);
            if (columns != null)
                columns.GetType().InvokeMember("AutoFit", IM, null, columns, null);
        }
        catch { }
    }

    /// <summary>
    /// Reads dimension and member IDs from the currently selected cell's comment.
    /// </summary>
    private (long dimId, long memberId) GetSelectedCellMember()
    {
        try
        {
            object? xlApp = ExcelDnaUtil.Application;
            if (xlApp == null) return (0, 0);

            var t = xlApp.GetType();
            object? activeCell = t.InvokeMember("ActiveCell", GP, null, xlApp, null);
            if (activeCell == null) return (0, 0);

            var cellType = activeCell.GetType();
            object? commentObj = cellType.InvokeMember("Comment", GP, null, activeCell, null);
            if (commentObj == null) return (0, 0);

            var commentType = commentObj.GetType();
            object? textObj = commentType.InvokeMember("Text", IM, null, commentObj, null);
            if (textObj == null) return (0, 0);

            var text = textObj.ToString() ?? "";
            var parts = text.Split('|');
            long dimId = 0, memberId = 0;
            foreach (var p in parts)
            {
                if (p.StartsWith("DIM:"))
                    long.TryParse(p[4..], out dimId);
                else if (p.StartsWith("MBR:"))
                    long.TryParse(p[4..], out memberId);
            }
            return (dimId, memberId);
        }
        catch { return (0, 0); }
    }


    private List<(long dimId, long memberId)> GetAllSelectedCellMembers()
    {
        var results = new List<(long dimId, long memberId)>();
        try
        {
            object? xlApp = ExcelDnaUtil.Application;
            if (xlApp == null) return results;
            var t = xlApp.GetType();
            object? selection = t.InvokeMember("Selection", GP, null, xlApp, null);
            if (selection == null) return results;
            var selType = selection.GetType();
            object? countObj = selType.InvokeMember("Count", GP, null, selection, null);
            int count = Convert.ToInt32(countObj ?? 1);
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    object? cell = selType.InvokeMember("Item", GP, null, selection, new object[] { i });
                    if (cell == null) continue;
                    var cellType = cell.GetType();
                    object? commentObj = cellType.InvokeMember("Comment", GP, null, cell, null);
                    if (commentObj == null) continue;
                    object? textObj = commentObj.GetType().InvokeMember("Text", IM, null, commentObj, null);
                    if (textObj == null) continue;
                    var text = textObj.ToString() ?? "";
                    var parts = text.Split('|');
                    long dimId = 0, memberId = 0;
                    foreach (var p in parts)
                    {
                        if (p.StartsWith("DIM:")) long.TryParse(p[4..], out dimId);
                        else if (p.StartsWith("MBR:")) long.TryParse(p[4..], out memberId);
                    }
                    if (dimId > 0 && memberId > 0)
                        results.Add((dimId, memberId));
                }
                catch { }
            }
        }
        catch { }
        return results;
    }
    #endregion

    #region Helpers

    private static void ShowMessage(string msg)
    {
        MessageBox.Show(msg, "MyOlap", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        var txt = new TextBox { Left = 16, Top = 50, Width = 400 };
        var ok = new Button { Text = "OK", Left = 220, Top = 100, Width = 100, Height = 36, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", Left = 330, Top = 100, Width = 100, Height = 36, DialogResult = DialogResult.Cancel };
        form.Controls.AddRange(new Control[] { lbl, txt, ok, cancel });
        form.AcceptButton = ok;
        form.CancelButton = cancel;
        var owner = new Win32Window(GetExcelHwnd());
        return form.ShowDialog(owner) == DialogResult.OK ? txt.Text : null;
    }

    private static IntPtr GetExcelHwnd()
    {
        try
        {
            object? xlApp = ExcelDnaUtil.Application;
            if (xlApp == null) return IntPtr.Zero;
            object? hwnd = xlApp.GetType().InvokeMember("Hwnd", GP, null, xlApp, null);
            if (hwnd == null) return IntPtr.Zero;
            return new IntPtr(Convert.ToInt32(hwnd));
        }
        catch
        {
            return IntPtr.Zero;
        }
    }

    #endregion
}

/// <summary>
/// IWin32Window wrapper to parent WinForms dialogs on the Excel window.
/// </summary>
internal class Win32Window : IWin32Window
{
    public IntPtr Handle { get; }
    public Win32Window(IntPtr handle) => Handle = handle;
}


