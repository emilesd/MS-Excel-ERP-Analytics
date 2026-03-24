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

    public override string GetCustomUI(string ribbonId)
    {
        return @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab id='tabMyOlap' label='MyOlap'>
        <group id='grpModel' label='Model'>
          <button id='btnSelectModel'  label='Select Model'  size='large' imageMso='OmsBusinessDataList'  onAction='OnSelectModel'/>
          <button id='btnRefreshData'  label='Refresh Data'  size='large' imageMso='Refresh'              onAction='OnRefreshData'/>
        </group>
        <group id='grpNavigate' label='Navigate'>
          <button id='btnPickMember'   label='Pick Member'   size='large' imageMso='ShowTreeView'         onAction='OnPickMember'/>
          <button id='btnDrillDown'    label='Drill Down'    size='normal' imageMso='ZoomIn'              onAction='OnDrillDown'/>
          <button id='btnDrillUp'      label='Drill Up'      size='normal' imageMso='ZoomOut'             onAction='OnDrillUp'/>
        </group>
        <group id='grpView' label='View'>
          <button id='btnSwapRowCol'   label='Swap to Row/Col' size='normal' imageMso='PivotTablePivot'   onAction='OnSwapRowCol'/>
          <button id='btnKeepSelected' label='Keep Selected' size='normal' imageMso='FilterBySelection'   onAction='OnKeepSelected'/>
          <button id='btnRemoveSelected' label='Remove Selected' size='normal' imageMso='Delete'          onAction='OnRemoveSelected'/>
          <button id='btnUndoLast'     label='Undo Last'     size='normal' imageMso='Undo'                onAction='OnUndoLast'/>
        </group>
        <group id='grpAdmin' label='Admin'>
          <button id='btnManageModel'  label='Manage Model'  size='large' imageMso='PropertySheet'        onAction='OnManageModel'/>
          <button id='btnLoadData'     label='Load Data'     size='normal' imageMso='ImportTextFile'       onAction='OnLoadData'/>
          <button id='btnSettings'     label='Settings'      size='normal' imageMso='ControlProperties'    onAction='OnSettings'/>
        </group>
        <group id='grpReport' label='Report'>
          <button id='btnExportPdf'    label='Export PDF'    size='large' imageMso='FileSaveAsPdfOrXps'    onAction='OnExportPdf'/>
        </group>
        <group id='grpInfo' label='Info'>
          <labelControl id='lblActiveModel' getLabel='GetActiveModelLabel'/>
          <labelControl id='lblVersion'     label='Version: v2.2'/>
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
                var modelId = mgr.CreateEmptyModel(name);
                _engine.SelectModel(modelId);
                WriteGridToSheet();
            }
            else if (form.SelectedModelId > 0)
            {
                _engine.SelectModel(form.SelectedModelId);
                WriteGridToSheet();
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
        WriteGridToSheet();
    }

    public void OnPickMember(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new MemberPickerForm(_engine.ActiveModel.Id);
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

        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0 || memberId == 0)
        { ShowMessage("Select a member cell (row or column header) to drill down on."); return; }

        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new DrillOptionsForm();
            if (form.ShowDialog(owner) != DialogResult.OK) return;

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

        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0 || memberId == 0)
        { ShowMessage("Select a member cell (row or column header) to drill up on."); return; }

        _engine.DrillUp(dimId, memberId);
        WriteGridToSheet();
    }

    public void OnSwapRowCol(IRibbonControl control)
    {
        if (_engine.ActiveModel == null) { ShowMessage("No model selected."); return; }
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
            form.ShowDialog(owner);
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
            form.ShowDialog(owner);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Load Data error:\n{ex.Message}", "MyOlap Error",
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

            if (grid.RowHeaders.Count == 0 && grid.ColHeaders.Count == 0)
            {
                ComSetCell(ws, 1, 1, $"Model '{modelName}' is ready.");
                ComSetCell(ws, 2, 1, "Use Manage Model to add dimensions/members, then Refresh Data.");
                return;
            }

            int headerRows = grid.ColDimensionNames.Count > 0
                ? grid.ColDimensionNames.Count + 1 : 1;
            int headerCols = grid.RowDimensionNames.Count > 0
                ? grid.RowDimensionNames.Count : 1;

            for (int i = 0; i < grid.RowDimensionNames.Count; i++)
                ComSetCell(ws, 1, i + 1, grid.RowDimensionNames[i]);

            for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
            {
                var combo = grid.ColHeaders[cIdx];
                for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                {
                    ComSetCell(ws, dIdx + 1, headerCols + cIdx + 1, grid.FormatMember(combo[dIdx]));
                    if (view != null && dIdx < view.ColAxes.Count)
                        ComSetComment(ws, dIdx + 1, headerCols + cIdx + 1,
                            $"DIM:{view.ColAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}");
                }
            }

            for (int rIdx = 0; rIdx < grid.RowHeaders.Count; rIdx++)
            {
                var combo = grid.RowHeaders[rIdx];
                for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                {
                    ComSetCell(ws, headerRows + rIdx + 1, dIdx + 1, grid.FormatMember(combo[dIdx]));
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
                        ComSetCell(ws, headerRows + rIdx + 1, headerCols + cIdx + 1, (double)val.Value);
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
            Text = title, Width = 400, Height = 200,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            MaximizeBox = false, MinimizeBox = false,
            AutoScaleMode = AutoScaleMode.Dpi
        };
        var lbl = new Label { Text = prompt, Left = 16, Top = 16, Width = 340 };
        var txt = new TextBox { Left = 16, Top = 46, Width = 340 };
        var ok = new Button { Text = "OK", Left = 180, Top = 90, Width = 85, Height = 32, DialogResult = DialogResult.OK };
        var cancel = new Button { Text = "Cancel", Left = 275, Top = 90, Width = 85, Height = 32, DialogResult = DialogResult.Cancel };
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
