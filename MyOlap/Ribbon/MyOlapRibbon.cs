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
    private readonly Dictionary<string, long> _sheetModels = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, (int rows, int cols, int povCols)> _sheetGridExtents = new(StringComparer.OrdinalIgnoreCase);
    // Stores headerRows/headerCols per sheet so the next render can locate header vs. data cells.
    private readonly Dictionary<string, (int headerRows, int headerCols)> _sheetHeaderLayout = new(StringComparer.OrdinalIgnoreCase);
    // Keyed by "DIM:X|MBR:Y"; values are user-typed note text snapshotted before each grid clear.
    private readonly Dictionary<string, string> _userNotesByMemberId = new();
    // row → member combo key ("|"-joined member IDs) saved AT render time; used by CaptureGridNotes.
    private readonly Dictionary<int, string> _prevRowComboKeys = new();
    // (row,col) → "DIM:X|MBR:Y" for row header cells, saved at render time.
    private readonly Dictionary<(int row, int col), string> _prevRowHeaderKeys = new();
    // (row,col) → "DIM:X|MBR:Y" for col header cells, saved at render time.
    private readonly Dictionary<(int row, int col), string> _prevColHeaderKeys = new();
    // Data cell notes captured before clear: (row,col) → note text; re-applied at shifted positions.
    private readonly Dictionary<(int row, int col), string> _prevDataCellNotes = new();
    private bool _excelEventsHooked;

    public void OnRibbonLoad(IRibbonUI ribbonUI)
    {
        _ribbonUi = ribbonUI;
        // Open the DB in the background so key/init cost is paid during Excel startup.
        System.Threading.Tasks.Task.Run(() =>
        {
            try { SqliteRepository.Instance.EnsureDatabaseCreated(); }
            catch { }
        });
        HookExcelEvents();
        // After Excel finishes opening the workbook, reconnect Info label from sheet metadata.
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SyncConnectionFromActiveSheet();
                RefreshInfoLabels();
            });
        }
        catch { }
    }

    private void HookExcelEvents()
    {
        if (_excelEventsHooked) return;
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            app.SheetActivate += (Action<object>)(_ => OnExcelContextChanged());
            app.WorkbookActivate += (Action<object>)(_ => OnExcelContextChanged());
            app.WorkbookOpen += (Action<object>)(_ => OnExcelContextChanged());
            _excelEventsHooked = true;
        }
        catch
        {
            // Dynamic COM event wire-up can fail on some hosts; QueueAsMacro + getLabel still work.
        }
    }

    private void OnExcelContextChanged()
    {
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SyncConnectionFromActiveSheet();
                RefreshInfoLabels();
            });
        }
        catch
        {
            try { SyncConnectionFromActiveSheet(); RefreshInfoLabels(); } catch { }
        }
    }

    /// <summary>
    /// Reconnects ActiveModel (and view when possible) from the active sheet's stored
    /// connection metadata so Info shows the source model after close/reopen.
    /// </summary>
    private void SyncConnectionFromActiveSheet()
    {
        try
        {
            var sheetName = GetActiveSheetName();
            long? modelId = _sheetModels.TryGetValue(sheetName, out var cached)
                ? cached
                : ReadStoredModelId() ?? ReadModelIdFromWorkbookMeta(sheetName);
            if (!modelId.HasValue) return;

            _sheetModels[sheetName] = modelId.Value;
            bool needView = _engine.ActiveModel?.Id != modelId.Value || _engine.CurrentView == null;
            if (!_engine.ActivateModel(modelId.Value)) return;
            if (!needView) return;

            var restored = ReadViewFromWorksheet(modelId.Value)
                ?? ReadViewFromWorkbookMeta(sheetName, modelId.Value);
            if (restored != null)
                _engine.SetCurrentView(restored);
            else
                _engine.TryLoadPersistedView(modelId.Value);
        }
        catch { }
    }

    private void RefreshInfoLabels()
    {
        try { _ribbonUi?.InvalidateControl("lblActiveModel"); } catch { }
    }

    /// <summary>Reconnects from sheet metadata when Excel was restarted with a saved report.</summary>
    private bool EnsureModelConnected()
    {
        if (_engine.ActiveModel == null)
            SyncConnectionFromActiveSheet();
        if (_engine.ActiveModel == null)
        {
            ShowMessage("No model selected.");
            return false;
        }
        return true;
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
        </group>
        <group id='grpNavigate' label='Navigate'>
          <button id='btnPickMember'   label='Pick Member'   size='large' imageMso='OrganizationChartInsert'         onAction='OnPickMember'/>
          <button id='btnDrillDown'    label='Drill Down'    size='normal' imageMso='OutlineShowDetail'              onAction='OnDrillDown'/>
          <button id='btnDrillUp'      label='Drill Up'      size='normal' imageMso='OutlineHideDetail'             onAction='OnDrillUp'/>
        </group>
        <group id='grpView' label='Layout'>
          <box id='boxViewLeft' boxStyle='vertical'>
            <button id='btnMoveToRow'    label='Move to Row'     size='normal' getImage='GetMoveToRowImage'    onAction='OnMoveToRow'/>
            <button id='btnMoveToCol'    label='Move to Column'  size='normal' getImage='GetMoveToColImage'    onAction='OnMoveToCol'/>
            <button id='btnMoveToHeader' label='Move to Header'  size='normal' getImage='GetMoveToHeaderImage' onAction='OnMoveToHeader'/>
          </box>
          <box id='boxViewRight' boxStyle='vertical'>
            <button id='btnKeepSelected'   label='Keep Selected'   size='normal' imageMso='FilterBySelection' onAction='OnKeepSelected'/>
            <button id='btnRemoveSelected' label='Remove Selected' size='normal' imageMso='Delete'            onAction='OnRemoveSelected'/>
            <button id='btnUndoLast'       label='Undo Last'       size='normal' imageMso='Undo'              onAction='OnUndoLast'/>
          </box>
        </group>
        <group id='grpAdmin' label='Admin'>
          <button id='btnManageModel'  label='Manage Model'  size='large'  imageMso='DesignMode'          onAction='OnManageModel'/>
          <button id='btnLoadData'     label='Load Data'     size='normal' imageMso='ImportTextFile'       onAction='OnLoadData'/>
          <button id='btnClearData'    label='Clear Data'    size='normal' imageMso='RecordsDeleteRecord'  onAction='OnClearData'/>
          <button id='btnSettings'     label='Settings'      size='normal' imageMso='ControlProperties'    onAction='OnSettings'/>
        </group>

        <group id='grpInfo' label='Info'>
          <labelControl id='lblActiveModel' getLabel='GetActiveModelLabel'/>
          <labelControl id='lblVersion'     label='Version: v1.5'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
    }

    public string GetActiveModelLabel(IRibbonControl control)
    {
        try
        {
            // Prefer the connection source stored on the active sheet (survives close/reopen).
            var sheetName = GetActiveSheetName();
            long? modelId = _sheetModels.TryGetValue(sheetName, out var cached)
                ? cached
                : ReadStoredModelId() ?? ReadModelIdFromWorkbookMeta(sheetName);
            if (modelId.HasValue)
            {
                var model = SqliteRepository.Instance.GetAllModels()
                    .FirstOrDefault(m => m.Id == modelId.Value);
                if (model != null)
                    return $"Model: {model.Name}";
            }
            var activeName = _engine.ActiveModel?.Name;
            if (!string.IsNullOrEmpty(activeName))
                return $"Model: {activeName}";
        }
        catch { }
        return "Model: (none)";
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
                _sheetModels[GetActiveSheetName()] = modelId;
                RefreshInfoLabels();
            }
            else if (form.SelectedModelId > 0)
            {
                var sheetName = GetActiveSheetName();
                long? prevModelId = _sheetModels.TryGetValue(sheetName, out var pm)
                    ? (long?)pm
                    : ReadStoredModelId();
                if (prevModelId.HasValue && prevModelId.Value != form.SelectedModelId)
                {
                    var prevName = SqliteRepository.Instance.GetAllModels().FirstOrDefault(m => m.Id == prevModelId.Value)?.Name ?? "Unknown";
                    var result = MessageBox.Show(
                        $"Current sheet connection source model ({prevName}) does not match currently selected model.\nDefault view will be used. Current values on the sheet will be cleared.\n\nDo you want to continue?",
                        "Model Mismatch", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.No) return;
                    ClearActiveSheet();
                    _engine.SelectModel(form.SelectedModelId);
                }
                else
                {
                    // Same model re-selected: keep the layout (do NOT reset to default).
                    // Restore order matches how the client keeps state durable across Excel restarts:
                    // 1) worksheet comment metadata (current grid on screen)
                    // 2) workbook meta sheet (saved with the .xlsx — respects Save / Don't Save)
                    // 3) in-memory session cache
                    // 4) SQLite ModelViews (same store as dimensions/settings)
                    _engine.SelectModel(form.SelectedModelId, preserveUndo: true);
                    var restored = ReadViewFromWorksheet(form.SelectedModelId)
                        ?? ReadViewFromWorkbookMeta(sheetName, form.SelectedModelId);
                    if (restored != null)
                        _engine.SetCurrentView(restored);
                    else if (!_engine.RestoreViewForSheet(sheetName, form.SelectedModelId))
                        _engine.TryLoadPersistedView(form.SelectedModelId);
                }
                WriteGridToSheet();
                _sheetModels[sheetName] = form.SelectedModelId;
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
        if (!EnsureModelConnected()) return;
        try
        {
            var (contextDimId, _) = GetSelectedCellMember();
            // Fallback: if user clicked a value cell (no DIM comment), use first row-axis dim then col-axis
            if (contextDimId == 0 && _engine.CurrentView != null)
            {
                if (_engine.CurrentView.RowAxes.Count > 0)
                    contextDimId = _engine.CurrentView.RowAxes[0].DimensionId;
                else if (_engine.CurrentView.ColAxes.Count > 0)
                    contextDimId = _engine.CurrentView.ColAxes[0].DimensionId;
            }
            // Collect members already on the axis for the context dimension so the dialog pre-populates them.
            var currentAxisMembers = new List<(long Id, string Name)>();
            bool dimIsOnRow = true;
            if (_engine.CurrentView != null && contextDimId > 0)
            {
                var rowAxis = _engine.CurrentView.RowAxes.FirstOrDefault(a => a.DimensionId == contextDimId);
                var colAxis = _engine.CurrentView.ColAxes.FirstOrDefault(a => a.DimensionId == contextDimId);
                var axis = rowAxis ?? colAxis;
                dimIsOnRow = colAxis == null; // on row unless it's on col axis
                if (axis != null)
                {
                    foreach (var mid in axis.VisibleMemberIds)
                    {
                        var m = SqliteRepository.Instance.GetMember(mid);
                        if (m != null) currentAxisMembers.Add((m.Id, m.DisplayName));
                    }
                }
            }
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new MemberPickerForm(_engine.ActiveModel.Id, contextDimId, currentAxisMembers, dimIsOnRow);
            if (form.ShowDialog(owner) != DialogResult.OK) return;

            var ids = form.SelectedMemberIds;
            if (ids.Count == 0) return;

            _engine.PickMembers(form.SelectedDimensionId, ids, form.PlaceOnRow);
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
        if (!EnsureModelConnected() || _engine.CurrentView == null)
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
            if (!WriteGridToSheet())
                _engine.Undo();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Drill Down error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnDrillUp(IRibbonControl control)
    {
        if (!EnsureModelConnected() || _engine.CurrentView == null)
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
        if (!WriteGridToSheet())
            _engine.Undo();
    }

    public System.Drawing.Bitmap GetMoveToRowImage(IRibbonControl _) => RibbonIcons.MoveToRow();
    public System.Drawing.Bitmap GetMoveToColImage(IRibbonControl _) => RibbonIcons.MoveToCol();
    public System.Drawing.Bitmap GetMoveToHeaderImage(IRibbonControl _) => RibbonIcons.MoveToHeader();

    public void OnMoveToRow(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        var (dimId, _) = GetSelectedCellMember();
        if (dimId == 0) { ShowMessage("Select a dimension or member cell first."); return; }
        var settings = SqliteRepository.Instance.GetSettings(_engine.ActiveModel.Id);
        if (settings.PreserveFormulas)
        {
            if (MessageBox.Show("Formulas and text in this worksheet will be lost. Do you want to continue?",
                "Move to Row", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
        }
        if (!_engine.MoveToRow(dimId)) { ShowMessage("This dimension is already on the row axis."); return; }
        if (!WriteGridToSheet()) _engine.Undo();
    }

    public void OnMoveToCol(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        var (dimId, _) = GetSelectedCellMember();
        if (dimId == 0) { ShowMessage("Select a dimension or member cell first."); return; }
        var settings = SqliteRepository.Instance.GetSettings(_engine.ActiveModel.Id);
        if (settings.PreserveFormulas)
        {
            if (MessageBox.Show("Formulas and text in this worksheet will be lost. Do you want to continue?",
                "Move to Column", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No) return;
        }
        if (!_engine.MoveToCol(dimId)) { ShowMessage("This dimension is already on the column axis."); return; }
        if (!WriteGridToSheet()) _engine.Undo();
    }

    public void OnMoveToHeader(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0) { ShowMessage("Select a dimension or member cell first."); return; }
        if (!_engine.MoveToHeader(dimId, memberId))
        {
            ShowMessage("Cannot move to header: the grid must keep at least one row dimension and one column dimension. Use Pick Member to rearrange dimensions first.");
            return;
        }
        if (!WriteGridToSheet())
            _engine.Undo();
    }

    public void OnKeepSelected(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0 || memberId == 0) { ShowMessage("Select a member cell first."); return; }
        _engine.KeepSelected(dimId, memberId);
        if (!WriteGridToSheet())
            _engine.Undo();
    }

    public void OnRemoveSelected(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        var (dimId, memberId) = GetSelectedCellMember();
        if (dimId == 0 || memberId == 0) { ShowMessage("Select a member cell first."); return; }
        _engine.RemoveSelected(dimId, memberId);
        if (!WriteGridToSheet())
            _engine.Undo();
    }

    public void OnUndoLast(IRibbonControl control)
    {
        if (!_engine.CanUndo)
        {
            if (_engine.UndoTotalPushed == 0)
                ShowMessage("Nothing to undo.");
            else if (_engine.UndoLimitReached)
                ShowMessage($"Undo limit reached – no more steps available (maximum {MyOlap.Core.UndoManager.MaxUndoLevels} steps).");
            else
                ShowMessage("No more steps to undo.");
            return;
        }
        _engine.Undo();
        WriteGridToSheet();
    }

    public void OnManageModel(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new ManageStructureForm(_engine.ActiveModel.Id);
            form.ShowDialog(owner);
            // Close with no edits: leave the sheet layout alone.
            // Structure changed (e.g. new dimension): keep current drill/axes, only sync
            // missing dims onto POV — do NOT reset to the model default view.
            if (form.StructureChanged)
            {
                if (_engine.CurrentView == null)
                    _engine.SelectModel(_engine.ActiveModel.Id);
                else
                    _engine.SyncViewWithModelStructure();
                WriteGridToSheet();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Manage Model error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnLoadData(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
        try
        {
            var owner = new Win32Window(GetExcelHwnd());
            using var form = new DataLoadForm(_engine.ActiveModel.Id);
            if (form.ShowDialog(owner) == DialogResult.OK)
            {
                _engine.InvalidateFactCache();
                WriteGridToSheet();
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Load Data error:\n{ex.Message}", "MyOlap Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnClearData(IRibbonControl control)
    {
        if (!EnsureModelConnected()) return;
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

            foreach (var m in viewMembers) cbView.Items.Add(m.DisplayName);
            foreach (var m in versionMembers) cbVersion.Items.Add(m.DisplayName);
            foreach (var m in yearMembers) cbYear.Items.Add(m.DisplayName);

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

            _engine.InvalidateFactCache();
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
        if (!EnsureModelConnected()) return;
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
        if (!EnsureModelConnected() || _engine.CurrentView == null)
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

    // Reads the text of a cell's existing comment/note. Returns "" if none.
    private static string ReadCellCommentText(object ws, int row, int col)
    {
        try
        {
            var wsType = ws.GetType();
            object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row, col });
            if (cell == null) return "";
            object? commentObj = cell.GetType().InvokeMember("Comment", GP, null, cell, null);
            if (commentObj == null) return "";
            object? textObj = commentObj.GetType().InvokeMember("Text", IM, null, commentObj, null);
            return textObj?.ToString() ?? "";
        }
        catch { return ""; }
    }

    // Returns everything after the first newline in a comment (the user note), or null if none.
    private static string? ExtractUserNote(string commentText)
    {
        int nl = commentText.IndexOf('\n');
        if (nl < 0) return null;
        string userPart = commentText[(nl + 1)..].Trim();
        return string.IsNullOrEmpty(userPart) ? null : userPart;
    }

    private static string? ExtractDimMbrKey(string commentText) => SheetMetaParser.ExtractDimMbrKey(commentText);

    private static long? ExtractModelIdFromComment(string commentText) => SheetMetaParser.ExtractModelId(commentText);

    // Captures user-added notes before the grid is cleared so they can be re-applied at the
    // correct member positions after the render, following any row/col shifts from drill-down.
    //
    // Header cells: user notes stored by "DIM:X|MBR:Y" key so they survive member position shifts.
    // Data cells:   notes stored by (row,col) and re-mapped via the row-shift mechanism.
    //               Data cell comments are EXPLICITLY CLEARED here so they don't remain at stale
    //               positions after ClearContents (which preserves comments).
    private void CaptureGridNotes(object ws, string sheetName)
    {
        _userNotesByMemberId.Clear();
        _prevDataCellNotes.Clear();
        // _prevRowComboKeys, _prevRowHeaderKeys, _prevColHeaderKeys are populated at render time
        // and must NOT be cleared here — they describe the grid currently on screen.

        if (!_sheetGridExtents.TryGetValue(sheetName, out var extents)) return;
        if (!_sheetHeaderLayout.TryGetValue(sheetName, out var layout)) return;

        var (prevR, prevC, _) = extents;
        var (prevHeaderRows, prevHeaderCols) = layout;
        if (prevR == 0) return;

        // ── Scan header cells for user notes ─────────────────────────────────────────────
        for (int row = 1; row <= prevR; row++)
        {
            for (int col = 1; col <= prevC; col++)
            {
                string text = ReadCellCommentText(ws, row, col);
                if (string.IsNullOrEmpty(text)) continue;

                string? key = ExtractDimMbrKey(text);
                if (key != null)
                {
                    // Standard format: "DIM:X|MBR:Y\nuser note"
                    string? userNote = ExtractUserNote(text);
                    if (userNote != null)
                        _userNotesByMemberId[key] = userNote;
                }
                else
                {
                    // Plain note — user replaced the metadata comment with their own text.
                    // Use position maps saved at the last render to infer the member key.
                    if (_prevRowHeaderKeys.TryGetValue((row, col), out var rk))
                        _userNotesByMemberId[rk] = text;
                    else if (_prevColHeaderKeys.TryGetValue((row, col), out var ck))
                        _userNotesByMemberId[ck] = text;
                }
            }
        }

        // ── Data cells: capture plain notes and clear them so they don't linger at stale
        //    positions after ClearContents (which preserves cell comments).
        for (int row = prevHeaderRows + 1; row <= prevR; row++)
        {
            for (int col = prevHeaderCols + 1; col <= prevC; col++)
            {
                string text = ReadCellCommentText(ws, row, col);
                if (string.IsNullOrEmpty(text)) continue;
                if (text.Contains("DIM:") || text.Contains("MBR:") || text.Contains("MODEL:")) continue;
                _prevDataCellNotes[(row, col)] = text;
                ClearCellComment(ws, row, col);
            }
        }
    }

    // Re-applies captured data-cell notes at their new row positions after a grid re-render.
    private void RestoreDataCellNotes(object ws, GridResult grid, int headerRows, int headerCols)
    {
        if (_prevDataCellNotes.Count == 0) return;

        // Map member combo key → new row in the freshly rendered grid
        var newRowByKey = new Dictionary<string, int>();
        for (int rIdx = 0; rIdx < grid.RowHeaders.Count; rIdx++)
        {
            var key = string.Join("|", grid.RowHeaders[rIdx].Select(m => m.Id.ToString()));
            newRowByKey.TryAdd(key, headerRows + rIdx + 1);
        }

        foreach (var ((oldRow, col), noteText) in _prevDataCellNotes)
        {
            if (!_prevRowComboKeys.TryGetValue(oldRow, out var comboKey)) continue;
            if (!newRowByKey.TryGetValue(comboKey, out var newRow)) continue;
            AddCellNote(ws, newRow, col, noteText);
        }
    }

    // Returns true = proceed with render; false = user chose to cancel.
    // Only called when PreserveFormulas is on. Skips the check on first render of a session
    // (prevR==0) because we have no record of what the grid owned previously.
    private bool CheckUserTextInNewGridArea(object ws, string sheetName,
        int newGridRows, int newGridCols, int newPovCols, int headerRows)
    {
        var (prevR, prevC, prevPov) = _sheetGridExtents.TryGetValue(sheetName, out var prev) ? prev : (0, 0, 0);
        if (prevR == 0) return true; // no previous grid known — nothing to warn about

        int clearCols = Math.Max(newGridCols, newPovCols);
        int prevClearCols = Math.Max(prevC, prevPov);
        var cells = new List<string>();

        void Check(int row, int col)
        {
            if (cells.Count >= 20) return;
            try
            {
                object? cell = ws.GetType().InvokeMember("Cells", GP, null, ws, new object[] { row, col });
                if (cell == null) return;
                object? val = cell.GetType().InvokeMember("Value2", GP, null, cell, null);
                if (val == null) return;
                if (val is string s && string.IsNullOrEmpty(s)) return;
                cells.Add($"{ColToLetter(col)}{row}");
            }
            catch { }
        }

        // New bottom rows (below old grid, within new grid width)
        for (int r = prevR + 1; r <= newGridRows && cells.Count < 20; r++)
            for (int c = 1; c <= newGridCols && cells.Count < 20; c++)
                Check(r, c);

        // New right columns in the data-row overlap area
        if (newGridCols > prevC)
            for (int r = headerRows + 1; r <= Math.Min(newGridRows, prevR) && cells.Count < 20; r++)
                for (int c = prevC + 1; c <= newGridCols && cells.Count < 20; c++)
                    Check(r, c);

        // New header-area columns (POV or col-header area grew)
        if (clearCols > prevClearCols && headerRows > 0)
            for (int r = 1; r <= headerRows && cells.Count < 20; r++)
                for (int c = prevClearCols + 1; c <= clearCols && cells.Count < 20; c++)
                    Check(r, c);

        if (cells.Count == 0) return true;

        var result = MessageBox.Show(
            "Formulas and Text in this worksheet will be lost. Do you want to continue?",
            "MyOlap – Text Will Be Wiped Out",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);
        return result == DialogResult.Yes;
    }

    private static string ColToLetter(int col)
    {
        string result = "";
        while (col > 0) { col--; result = (char)('A' + col % 26) + result; col /= 26; }
        return result;
    }

    private static void ClearCellComment(object ws, int row, int col)
    {
        try
        {
            var wsType = ws.GetType();
            object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row, col });
            cell?.GetType().InvokeMember("ClearComments", IM, null, cell, null);
        }
        catch { }
    }

    private static void AddCellNote(object ws, int row, int col, string text)
    {
        try
        {
            var wsType = ws.GetType();
            object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row, col });
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

    // Writes a 2D array to a range using per-cell scalar writes via the range's own Cells(r,c) indexer.
    // Range.Value2 = 2D array fails intermittently via COM after ClearContents regardless of SAFEARRAY
    // lower bounds. Per-cell scalar writes are always safe; with ScreenUpdating=false the overhead is
    // acceptable for typical OLAP grid sizes.
    private static void SetRangeValue2(object rng, object[,] arr)
    {
        int rows = arr.GetLength(0);
        int cols = arr.GetLength(1);
        var rngType = rng.GetType();
        for (int r = 1; r <= rows; r++)
        {
            for (int c = 1; c <= cols; c++)
            {
                try
                {
                    object? cell = rngType.InvokeMember("Cells", GP, null, rng, new object[] { r, c });
                    if (cell == null) continue;
                    try { cell.GetType().InvokeMember("UnMerge", IM, null, cell, null); } catch { }
                    cell.GetType().InvokeMember("Value2", SP, null, cell, new object[] { arr[r - 1, c - 1] ?? "" });
                }
                catch { }
            }
        }
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

    private static void ComClearRange(object ws, int startRow, int startCol, int endRow, int endCol)
    {
        try
        {
            var wsType = ws.GetType();
            object? topLeft = wsType.InvokeMember("Cells", GP, null, ws, new object[] { startRow, startCol });
            object? bottomRight = wsType.InvokeMember("Cells", GP, null, ws, new object[] { endRow, endCol });
            if (topLeft == null || bottomRight == null) return;
            object? range = wsType.InvokeMember("Range", GP, null, ws, new object[] { topLeft, bottomRight });
            if (range == null) return;
            // UnMerge() handles partial merges (merge area straddling the range boundary);
            // MergeCells=false silently fails on partial merges leaving cells still merged.
            try { range.GetType().InvokeMember("UnMerge", IM, null, range, null); } catch { }
            try { range.GetType().InvokeMember("ClearContents", IM, null, range, null); } catch { }
        }
        catch { }
    }

    // Write a cell by cell within a 1-row band — guaranteed scalar, never fails on array/merge edge cases.
    private static void WriteRowCells(object ws, int row, int startCol, object?[] values)
    {
        var wsType = ws.GetType();
        for (int i = 0; i < values.Length; i++)
        {
            try
            {
                object? cell = wsType.InvokeMember("Cells", GP, null, ws, new object[] { row, startCol + i });
                if (cell == null) continue;
                try { cell.GetType().InvokeMember("UnMerge", IM, null, cell, null); } catch { }
                cell.GetType().InvokeMember("Value2", SP, null, cell, new object[] { values[i] ?? "" });
            }
            catch { }
        }
    }

    private static void ClearActiveSheet()
    {
        try
        {
            object? ws = GetActiveSheet();
            if (ws == null) return;
            object? usedRange = ws.GetType().InvokeMember("UsedRange", GP, null, ws, null);
            if (usedRange != null)
                usedRange.GetType().InvokeMember("Clear", IM, null, usedRange, null);
        }
        catch { }
    }

    private static (int rows, int cols) GetUsedRangeExtent(object ws)
    {
        try
        {
            var wsType = ws.GetType();
            object? usedRange = wsType.InvokeMember("UsedRange", GP, null, ws, null);
            if (usedRange == null) return (0, 0);
            var urType = usedRange.GetType();
            object? rowsObj = urType.InvokeMember("Rows", GP, null, usedRange, null);
            object? colsObj = urType.InvokeMember("Columns", GP, null, usedRange, null);
            int r = Convert.ToInt32(rowsObj?.GetType().InvokeMember("Count", GP, null, rowsObj, null) ?? 0);
            int c = Convert.ToInt32(colsObj?.GetType().InvokeMember("Count", GP, null, colsObj, null) ?? 0);
            return (r, c);
        }
        catch { return (0, 0); }
    }

    private long? ReadStoredModelId()
    {
        try
        {
            object? ws = GetActiveSheet();
            if (ws == null) return null;
            // Prefer A1 (where we always write MODEL:), then scan all sheet comments.
            var a1 = ExtractModelIdFromComment(ReadCellCommentText(ws, 1, 1));
            if (a1.HasValue) return a1;
            foreach (var (_, _, text) in EnumerateSheetComments(ws))
            {
                var id = ExtractModelIdFromComment(text);
                if (id.HasValue) return id;
            }
        }
        catch { }
        return null;
    }

    // Enumerates worksheet Notes/Comments via the Comments collection (indexed Item access).
    // Avoids Range.SpecialCells + foreach — that path fails silently with Excel COM/reflection
    // (IEnumerable cast), which made same-session restore appear to work only because the
    // in-memory _sheetViews cache was still warm.
    private static List<(int row, int col, string text)> EnumerateSheetComments(object ws)
    {
        var result = new List<(int row, int col, string text)>();
        try
        {
            object? comments = ws.GetType().InvokeMember("Comments", GP, null, ws, null);
            if (comments == null) return result;
            var commentsType = comments.GetType();
            int count = Convert.ToInt32(commentsType.InvokeMember("Count", GP, null, comments, null) ?? 0);
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    object? comment = commentsType.InvokeMember("Item", GP, null, comments, new object[] { i });
                    if (comment == null) continue;
                    var commentType = comment.GetType();
                    object? parent = commentType.InvokeMember("Parent", GP, null, comment, null);
                    if (parent == null) continue;
                    int r = Convert.ToInt32(parent.GetType().InvokeMember("Row", GP, null, parent, null));
                    int c = Convert.ToInt32(parent.GetType().InvokeMember("Column", GP, null, parent, null));
                    object? textObj = commentType.InvokeMember("Text", IM, null, comment, null);
                    result.Add((r, c, textObj?.ToString() ?? ""));
                }
                catch { }
            }
        }
        catch { }
        return result;
    }

    // Reads the header metadata comments (DIM:x|MBR:y) from the active sheet and rebuilds the
    // view layout from them. Returns null when the sheet has no intact MyOlap grid, so callers
    // can fall back to a saved or default view.
    private ViewState? ReadViewFromWorksheet(long modelId)
    {
        try
        {
            object? ws = GetActiveSheet();
            if (ws == null) return null;

            var cells = new List<(int row, int col, long dimId, long mbrId)>();
            foreach (var (r, c, text) in EnumerateSheetComments(ws))
            {
                var key = ExtractDimMbrKey(text);
                if (key != null && SheetMetaParser.TryParseDimMbr(key, out var dimId, out var mbrId))
                    cells.Add((r, c, dimId, mbrId));
            }
            if (cells.Count == 0) return null;
            return _engine.BuildViewFromMetadata(modelId, cells);
        }
        catch { return null; }
    }

    private const string MetaSheetName = "_MyOlapMeta";
    private const int XlSheetVeryHidden = 2;

    private static object? GetWorkbookFromSheet(object ws)
    {
        try { return ws.GetType().InvokeMember("Parent", GP, null, ws, null); }
        catch { return null; }
    }

    private static object? GetOrCreateMetaSheet(object workbook)
    {
        var wbType = workbook.GetType();
        object? sheets = wbType.InvokeMember("Worksheets", GP, null, workbook, null);
        if (sheets == null) return null;
        var sheetsType = sheets.GetType();
        try
        {
            return sheetsType.InvokeMember("Item", GP, null, sheets, new object[] { MetaSheetName });
        }
        catch
        {
            // Create very-hidden meta sheet (travels with Save / Don't Save like any worksheet).
            object? added = sheetsType.InvokeMember("Add", IM, null, sheets, null);
            if (added == null) return null;
            added.GetType().InvokeMember("Name", SP, null, added, new object[] { MetaSheetName });
            try { added.GetType().InvokeMember("Visible", SP, null, added, new object[] { XlSheetVeryHidden }); } catch { }
            return added;
        }
    }

    private static void WriteViewToWorkbookMeta(object ws, string sheetName, ViewState? view)
    {
        if (view == null || string.IsNullOrEmpty(sheetName)) return;
        object? wb = GetWorkbookFromSheet(ws);
        if (wb == null) return;
        object? meta = GetOrCreateMetaSheet(wb);
        if (meta == null) return;

        string payload = ViewStateCodec.Serialize(view);
        var metaType = meta.GetType();
        // Find existing row for this sheet name in column A; otherwise append.
        int row = 1;
        int emptyRow = 0;
        for (; row <= 500; row++)
        {
            object? cell = metaType.InvokeMember("Cells", GP, null, meta, new object[] { row, 1 });
            object? val = cell?.GetType().InvokeMember("Value2", GP, null, cell, null);
            var name = val?.ToString() ?? "";
            if (string.IsNullOrEmpty(name)) { if (emptyRow == 0) emptyRow = row; break; }
            if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                object? b = metaType.InvokeMember("Cells", GP, null, meta, new object[] { row, 2 });
                b?.GetType().InvokeMember("Value2", SP, null, b, new object[] { payload });
                return;
            }
        }
        int target = emptyRow > 0 ? emptyRow : row;
        object? aCell = metaType.InvokeMember("Cells", GP, null, meta, new object[] { target, 1 });
        object? bCell = metaType.InvokeMember("Cells", GP, null, meta, new object[] { target, 2 });
        aCell?.GetType().InvokeMember("Value2", SP, null, aCell, new object[] { sheetName });
        bCell?.GetType().InvokeMember("Value2", SP, null, bCell, new object[] { payload });
    }

    private string? ReadWorkbookMetaPayload(string sheetName)
    {
        try
        {
            object? ws = GetActiveSheet();
            if (ws == null) return null;
            object? wb = GetWorkbookFromSheet(ws);
            if (wb == null) return null;
            object? sheets = wb.GetType().InvokeMember("Worksheets", GP, null, wb, null);
            if (sheets == null) return null;
            object? meta;
            try { meta = sheets.GetType().InvokeMember("Item", GP, null, sheets, new object[] { MetaSheetName }); }
            catch { return null; }
            if (meta == null) return null;

            var metaType = meta.GetType();
            for (int row = 1; row <= 500; row++)
            {
                object? cell = metaType.InvokeMember("Cells", GP, null, meta, new object[] { row, 1 });
                object? val = cell?.GetType().InvokeMember("Value2", GP, null, cell, null);
                var name = val?.ToString() ?? "";
                if (string.IsNullOrEmpty(name)) break;
                if (!string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase)) continue;

                object? b = metaType.InvokeMember("Cells", GP, null, meta, new object[] { row, 2 });
                object? payloadObj = b?.GetType().InvokeMember("Value2", GP, null, b, null);
                return payloadObj?.ToString();
            }
        }
        catch { }
        return null;
    }

    private long? ReadModelIdFromWorkbookMeta(string sheetName)
        => ViewStateCodec.TryParseModelId(ReadWorkbookMetaPayload(sheetName));

    private ViewState? ReadViewFromWorkbookMeta(string sheetName, long modelId)
    {
        try
        {
            var payload = ReadWorkbookMetaPayload(sheetName);
            var view = ViewStateCodec.Deserialize(payload, id =>
                SqliteRepository.Instance.GetDimensions(modelId).FirstOrDefault(d => d.Id == id)?.Name);
            return view != null && view.ModelId == modelId ? view : null;
        }
        catch { return null; }
    }

    /// <summary>
    /// Writes the OLAP grid to the active worksheet using COM via reflection.
    /// No dynamic keyword, no C API - purely reflection-based COM calls.
    /// </summary>
    private bool WriteGridToSheet()
    {
        object? xlApp = null;
        string _step = "init";
        try
        {
            var grid = _engine.BuildGrid();
            var view = _engine.CurrentView;
            var modelName = _engine.ActiveModel?.Name ?? "Model";

            object? ws = GetActiveSheet();
            if (ws == null)
            {
                ShowMessage("No active worksheet found. Please open or create a worksheet first.");
                return true;
            }

            var sheetName = GetActiveSheetName();
            var modelId = _engine.ActiveModel?.Id ?? 0;
            var settings = _engine.ActiveModel != null
                ? SqliteRepository.Instance.GetSettings(_engine.ActiveModel.Id)
                : new ModelSettings();

            int povRowCount = (view != null && view.PovSelections.Count > 0) ? 2 : 0;
            int headerRows = (grid.ColDimensionNames.Count > 0
                ? grid.ColDimensionNames.Count + 1 : 1) + povRowCount;
            int headerCols = grid.RowDimensionNames.Count > 0
                ? grid.RowDimensionNames.Count : 1;
            int newGridRows = headerRows + grid.RowHeaders.Count;
            int newGridCols = headerCols + (grid.ColHeaders.Count > 0 ? grid.ColHeaders.Count : 0);
            int newPovCols = view?.PovSelections.Count ?? 0;

            // Before touching the sheet: warn the user if text they placed will be wiped out
            // by the grid expanding into those cells. Do this before ScreenUpdating is off
            // so Excel is fully responsive during the dialog.
            if (settings.PreserveFormulas && !CheckUserTextInNewGridArea(ws, sheetName, newGridRows, newGridCols, newPovCols, headerRows))
                return false;

            // Save view only after user confirms (or no warning needed) so a cancelled render
            // doesn't leave the sheet's stored view pointing at the post-operation state.
            try { _engine.SaveViewForSheet(sheetName); } catch { }
            // Durable layout: SQLite (cross-session) + workbook meta sheet (Save/Don't Save).
            try { _engine.PersistCurrentView(); } catch { }
            try { WriteViewToWorkbookMeta(ws, sheetName, view); } catch { }

            // Suspend screen refresh and formula recalculation for the entire render.
            // This eliminates per-cell flicker and is the single biggest speed improvement.
            xlApp = ExcelDnaUtil.Application;
            if (xlApp != null)
            {
                try { xlApp.GetType().InvokeMember("ScreenUpdating", SP, null, xlApp, new object[] { false }); } catch { }
                try { xlApp.GetType().InvokeMember("Calculation", SP, null, xlApp, new object[] { -4135 }); } catch { } // xlCalculationManual
            }

            // Capture user-added notes before the grid is cleared.  Populates _userNotesByMemberId
            // (for header-cell notes) and _prevDataCellNotes / _prevRowComboKeys (for data-cell
            // notes that need to follow member rows on drill-down).
            CaptureGridNotes(ws, sheetName);

            if (settings.PreserveFormulas)
            {
                var (prevR, prevC, prevPov) = _sheetGridExtents.TryGetValue(sheetName, out var prev) ? prev : (0, 0, 0);
                int clearCols = Math.Max(newGridCols, newPovCols);

                // Header rows (POV + col header labels): clear at full POV width.
                if (headerRows > 0 && clearCols > 0)
                    ComClearRange(ws, 1, 1, headerRows, clearCols);

                // Data rows: clear ONLY the actual grid columns (1..newGridCols).
                // POV columns (beyond newGridCols) in data rows contain no system content,
                // so wiping them at clearCols width would destroy user-placed text/formulas
                // (e.g. "hi" in col D when the grid is only 2 cols wide).
                if (newGridRows > headerRows && newGridCols > 0)
                    ComClearRange(ws, headerRows + 1, 1, newGridRows, newGridCols);

                // Stale tail rows that fell off the bottom after drill-up or model reset.
                int staleCols = Math.Max(newGridCols, prevC);
                if (prevR > newGridRows && staleCols > 0)
                    ComClearRange(ws, newGridRows + 1, 1, prevR, staleCols);

                // Stale data columns in the data-row area (e.g. Time drill-up collapsed cols).
                if (prevC > newGridCols && prevR > headerRows)
                    ComClearRange(ws, headerRows + 1, newGridCols + 1, Math.Min(newGridRows, prevR), prevC);

                // Stale header-area columns (old POV/col-header cols that no longer exist).
                int prevClearCols = Math.Max(prevC, prevPov);
                if (prevClearCols > clearCols && headerRows > 0)
                    ComClearRange(ws, 1, clearCols + 1, headerRows, prevClearCols);
            }
            else
            {
                ComClearSheet(ws);
            }
            _sheetGridExtents[sheetName] = (newGridRows, newGridCols, newPovCols);
            _sheetHeaderLayout[sheetName] = (headerRows, headerCols);

            try
            {
                if (xlApp != null)
                    xlApp.GetType().InvokeMember("DisplayCommentIndicator", SP, null, xlApp, new object[] { 0 });
            }
            catch { }

            if (grid.RowHeaders.Count == 0 && grid.ColHeaders.Count == 0)
            {
                ComSetCell(ws, 1, 1, $"Model '{modelName}' is ready.");
                ComSetComment(ws, 1, 1, $"MODEL:{modelId}");
                ComSetCell(ws, 2, 1, "Use Manage Model to add dimensions/members.");
                return true;
            }

            if (view != null && view.PovSelections.Count > 0)
            {
                var allDims = SqliteRepository.Instance.GetDimensions(view.ModelId);
                // Collect valid (dim, member) pairs first so we know the exact count.
                var povPairs = new List<(Dimension dim, Member member)>();
                foreach (var kvp in view.PovSelections)
                {
                    var dim = allDims.FirstOrDefault(d => d.Id == kvp.Key);
                    var member = SqliteRepository.Instance.GetMember(kvp.Value);
                    if (dim != null && member != null)
                        povPairs.Add((dim, member));
                }
                if (povPairs.Count > 0)
                {
                    _step = "POV-write";
                    var wsType0 = ws.GetType();
                    // Write each POV label as an individual scalar — avoids all array/merge edge cases.
                    // POV is at most a handful of cells so per-cell writes have no perf cost.
                    var povLabels = new object?[povPairs.Count];
                    for (int i = 0; i < povPairs.Count; i++)
                        povLabels[i] = $"{povPairs[i].dim.Name}: {grid.FormatMember(povPairs[i].member)}";
                    WriteRowCells(ws, 1, 1, povLabels);
                    // Bold the POV row as one range operation after all values are written.
                    object? povTl = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { 1, 1 });
                    object? povBr = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { 1, povPairs.Count });
                    object? povRng = wsType0.InvokeMember("Range", GP, null, ws, new object[] { povTl, povBr });
                    if (povRng != null)
                    {
                        var povFont = povRng.GetType().InvokeMember("Font", GP, null, povRng, null);
                        try { povFont?.GetType().InvokeMember("Bold", SP, null, povFont, new object[] { true }); } catch { }
                    }
                    // Comments still need individual calls; first cell carries the model marker.
                    for (int i = 0; i < povPairs.Count; i++)
                    {
                        var metaText = i == 0
                            ? $"MODEL:{modelId}|DIM:{povPairs[i].dim.Id}|MBR:{povPairs[i].member.Id}"
                            : $"DIM:{povPairs[i].dim.Id}|MBR:{povPairs[i].member.Id}";
                        var dimMbrKey = $"DIM:{povPairs[i].dim.Id}|MBR:{povPairs[i].member.Id}";
                        if (_userNotesByMemberId.TryGetValue(dimMbrKey, out var pn))
                            metaText = $"{metaText}\n{pn}";
                        ComSetComment(ws, 1, i + 1, metaText);
                    }
                }
            }
            else
            {
                // No POV — write model marker to A1 so it persists across sessions
                ComSetComment(ws, 1, 1, $"MODEL:{modelId}");
            }

            // ── Row dimension name labels (one batch write + one bold range) ──────────────
            if (grid.RowDimensionNames.Count > 0)
            {
                _step = "RowDimNames-write";
                var wsType0 = ws.GetType();
                var rowDimArr = new object[1, grid.RowDimensionNames.Count];
                for (int i = 0; i < grid.RowDimensionNames.Count; i++)
                    rowDimArr[0, i] = grid.RowDimensionNames[i];
                object? rdTl = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { 1 + povRowCount, 1 });
                object? rdBr = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { 1 + povRowCount, grid.RowDimensionNames.Count });
                object? rdRng = wsType0.InvokeMember("Range", GP, null, ws, new object[] { rdTl, rdBr });
                if (rdRng != null)
                {
                    try { rdRng.GetType().InvokeMember("UnMerge", IM, null, rdRng, null); } catch { }
                    SetRangeValue2(rdRng, rowDimArr);
                    var rdFont = rdRng.GetType().InvokeMember("Font", GP, null, rdRng, null);
                    try { rdFont?.GetType().InvokeMember("Bold", SP, null, rdFont, new object[] { true }); } catch { }
                }
            }

            // ── Column headers: batch text + bold, then individual comments ──────────────
            if (grid.ColHeaders.Count > 0 && grid.ColHeaders[0].Count > 0)
            {
                _step = "ColHeaders-write";
                int colDimCount = grid.ColHeaders[0].Count;
                var wsType0 = ws.GetType();
                var colHdrArr = new object[colDimCount, grid.ColHeaders.Count];
                for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
                {
                    var combo = grid.ColHeaders[cIdx];
                    for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                        colHdrArr[dIdx, cIdx] = grid.FormatMember(combo[dIdx]) ?? "";
                }
                object? chTl = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { 1 + povRowCount, headerCols + 1 });
                object? chBr = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { povRowCount + colDimCount, headerCols + grid.ColHeaders.Count });
                object? chRng = wsType0.InvokeMember("Range", GP, null, ws, new object[] { chTl, chBr });
                if (chRng != null)
                {
                    try { chRng.GetType().InvokeMember("UnMerge", IM, null, chRng, null); } catch { }
                    SetRangeValue2(chRng, colHdrArr);
                    // Clear all bold first so previous rollup formatting doesn't linger.
                    var chFont = chRng.GetType().InvokeMember("Font", GP, null, chRng, null);
                    try { chFont?.GetType().InvokeMember("Bold", SP, null, chFont, new object[] { false }); } catch { }
                }
                // Bold only parent (roll-up) members in column headers — same rule as row headers.
                for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
                {
                    var combo = grid.ColHeaders[cIdx];
                    for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                        if (SqliteRepository.Instance.GetChildren(combo[dIdx].Id).Count > 0)
                            ComSetBold(ws, 1 + povRowCount + dIdx, headerCols + cIdx + 1, true);
                }
                // Comments — write system metadata. Col position map updated after RestoreDataCellNotes.
                if (view != null)
                {
                    for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
                    {
                        var combo = grid.ColHeaders[cIdx];
                        for (int dIdx = 0; dIdx < combo.Count && dIdx < view.ColAxes.Count; dIdx++)
                        {
                            var dimMbrKey = $"DIM:{view.ColAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}";
                            var commentText = dimMbrKey;
                            if (_userNotesByMemberId.TryGetValue(dimMbrKey, out var cn))
                                commentText = $"{dimMbrKey}\n{cn}";
                            ComSetComment(ws, dIdx + 1 + povRowCount, headerCols + cIdx + 1, commentText);
                        }
                    }
                }
            }

            // ── Row headers: batch text, selective bold, then individual comments ─────────
            if (grid.RowHeaders.Count > 0 && headerCols > 0)
            {
                _step = "RowHeaders-write";
                int rCount = grid.RowHeaders.Count;
                var wsType0 = ws.GetType();
                var rowHdrArr = new object[rCount, headerCols];
                for (int rIdx = 0; rIdx < rCount; rIdx++)
                {
                    var combo = grid.RowHeaders[rIdx];
                    for (int dIdx = 0; dIdx < combo.Count && dIdx < headerCols; dIdx++)
                        rowHdrArr[rIdx, dIdx] = grid.FormatMember(combo[dIdx]) ?? "";
                }
                object? rhTl = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { headerRows + 1, 1 });
                object? rhBr = wsType0.InvokeMember("Cells", GP, null, ws, new object[] { headerRows + rCount, headerCols });
                object? rhRng = wsType0.InvokeMember("Range", GP, null, ws, new object[] { rhTl, rhBr });
                if (rhRng != null)
                {
                    try { rhRng.GetType().InvokeMember("UnMerge", IM, null, rhRng, null); } catch { }
                    SetRangeValue2(rhRng, rowHdrArr);
                    // Clear all bold first so previous rollup formatting doesn't linger.
                    var rhFont = rhRng.GetType().InvokeMember("Font", GP, null, rhRng, null);
                    try { rhFont?.GetType().InvokeMember("Bold", SP, null, rhFont, new object[] { false }); } catch { }
                }

                // Bold only parent members (has children) — individual calls are fast with ScreenUpdating=false.
                for (int rIdx = 0; rIdx < rCount; rIdx++)
                {
                    var combo = grid.RowHeaders[rIdx];
                    for (int dIdx = 0; dIdx < combo.Count && dIdx < headerCols; dIdx++)
                        if (SqliteRepository.Instance.GetChildren(combo[dIdx].Id).Count > 0)
                            ComSetBold(ws, headerRows + rIdx + 1, dIdx + 1, true);
                }
                // Comments — write system metadata. Position maps are updated after RestoreDataCellNotes.
                if (view != null)
                {
                    for (int rIdx = 0; rIdx < rCount; rIdx++)
                    {
                        var combo = grid.RowHeaders[rIdx];
                        for (int dIdx = 0; dIdx < combo.Count && dIdx < view.RowAxes.Count; dIdx++)
                        {
                            var dimMbrKey = $"DIM:{view.RowAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}";
                            var commentText = dimMbrKey;
                            if (_userNotesByMemberId.TryGetValue(dimMbrKey, out var rn))
                                commentText = $"{dimMbrKey}\n{rn}";
                            ComSetComment(ws, headerRows + rIdx + 1, dIdx + 1, commentText);
                        }
                    }
                }
            }

            // Write all data values in one Range.Value2 assignment instead of cell-by-cell.
            if (grid.RowHeaders.Count > 0 && grid.ColHeaders.Count > 0)
            {
                _step = "DataValues-write";
                int dRows = grid.RowHeaders.Count;
                int dCols = grid.ColHeaders.Count;
                object[,] dataArr = new object[dRows, dCols];
                for (int rIdx = 0; rIdx < dRows; rIdx++)
                    for (int cIdx = 0; cIdx < dCols; cIdx++)
                        dataArr[rIdx, cIdx] = grid.Values[rIdx, cIdx].HasValue
                            ? (object)(double)grid.Values[rIdx, cIdx]!.Value
                            : null;  // null → VT_EMPTY (blank); DBNull → VT_NULL which Excel rejects

                var wsType = ws.GetType();
                object? tl = wsType.InvokeMember("Cells", GP, null, ws, new object[] { headerRows + 1, headerCols + 1 });
                object? br = wsType.InvokeMember("Cells", GP, null, ws, new object[] { headerRows + dRows, headerCols + dCols });
                if (tl != null && br != null)
                {
                    object? dataRange = wsType.InvokeMember("Range", GP, null, ws, new object[] { tl, br });
                    if (dataRange != null)
                    {
                        try { dataRange.GetType().InvokeMember("UnMerge", IM, null, dataRange, null); } catch { }
                        SetRangeValue2(dataRange, dataArr);
                        try { dataRange.GetType().InvokeMember("NumberFormat", SP, null, dataRange, new object[] { "#,##0.00" }); } catch { }
                        // Clear all bold first, then bold the entire row of data for each rollup row.
                        var dataFont = dataRange.GetType().InvokeMember("Font", GP, null, dataRange, null);
                        try { dataFont?.GetType().InvokeMember("Bold", SP, null, dataFont, new object[] { false }); } catch { }
                        for (int rIdx = 0; rIdx < dRows; rIdx++)
                        {
                            var combo = grid.RowHeaders[rIdx];
                            bool isRollup = combo.Any(m => SqliteRepository.Instance.GetChildren(m.Id).Count > 0);
                            if (!isRollup) continue;
                            object? rowTl = wsType.InvokeMember("Cells", GP, null, ws, new object[] { headerRows + rIdx + 1, headerCols + 1 });
                            object? rowBr = wsType.InvokeMember("Cells", GP, null, ws, new object[] { headerRows + rIdx + 1, headerCols + dCols });
                            object? rowRng = wsType.InvokeMember("Range", GP, null, ws, new object[] { rowTl, rowBr });
                            if (rowRng != null)
                            {
                                var rowFont = rowRng.GetType().InvokeMember("Font", GP, null, rowRng, null);
                                try { rowFont?.GetType().InvokeMember("Bold", SP, null, rowFont, new object[] { true }); } catch { }
                            }
                        }
                    }
                }
            }

            // AutoFit used columns while ScreenUpdating is still off — no per-cell flicker,
            // widths apply atomically when ScreenUpdating restores to true in the finally block.
            try
            {
                object? usedRange = ws.GetType().InvokeMember("UsedRange", GP, null, ws, null);
                if (usedRange != null)
                {
                    object? cols = usedRange.GetType().InvokeMember("Columns", GP, null, usedRange, null);
                    cols?.GetType().InvokeMember("AutoFit", IM, null, cols, null);
                }
            }
            catch { }

            // Re-apply data-cell notes that were shifted by drill-down row changes.
            RestoreDataCellNotes(ws, grid, headerRows, headerCols);

            // Snapshot position maps NOW — after RestoreDataCellNotes has finished using the
            // OLD maps. These become the "previous render" maps for the next CaptureGridNotes call.
            _prevRowComboKeys.Clear();
            _prevRowHeaderKeys.Clear();
            _prevColHeaderKeys.Clear();
            if (view != null)
            {
                for (int rIdx = 0; rIdx < grid.RowHeaders.Count; rIdx++)
                {
                    var combo = grid.RowHeaders[rIdx];
                    int rowNum = headerRows + rIdx + 1;
                    var sb2 = new System.Text.StringBuilder();
                    for (int dIdx = 0; dIdx < combo.Count && dIdx < view.RowAxes.Count; dIdx++)
                    {
                        _prevRowHeaderKeys[(rowNum, dIdx + 1)] = $"DIM:{view.RowAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}";
                        if (sb2.Length > 0) sb2.Append('|');
                        sb2.Append(combo[dIdx].Id);
                    }
                    if (sb2.Length > 0)
                        _prevRowComboKeys[rowNum] = sb2.ToString();
                }
                for (int cIdx = 0; cIdx < grid.ColHeaders.Count; cIdx++)
                {
                    var combo = grid.ColHeaders[cIdx];
                    int colNum = headerCols + cIdx + 1;
                    for (int dIdx = 0; dIdx < combo.Count && dIdx < view.ColAxes.Count; dIdx++)
                        _prevColHeaderKeys[(dIdx + 1 + povRowCount, colNum)] = $"DIM:{view.ColAxes[dIdx].DimensionId}|MBR:{combo[dIdx].Id}";
                }
            }

        }
        catch (Exception ex)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Error rendering grid [step={_step}]: {ex.Message}");
            var inner = ex.InnerException;
            int depth = 0;
            while (inner != null && depth++ < 4)
            {
                sb.AppendLine($"\nCaused by: {inner.Message}");
                inner = inner.InnerException;
            }
            sb.AppendLine($"\nStack: {ex.StackTrace}");
            ShowMessage(sb.ToString());
        }
        finally
        {
            // Always restore screen updating and auto-calculation, even on error.
            if (xlApp != null)
            {
                try { xlApp.GetType().InvokeMember("Calculation", SP, null, xlApp, new object[] { -4105 }); } catch { } // xlCalculationAutomatic
                try { xlApp.GetType().InvokeMember("ScreenUpdating", SP, null, xlApp, new object[] { true }); } catch { }
            }
        }
        return true;
    }

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
            return SheetMetaParser.TryParseDimMbr(SheetMetaParser.ExtractDimMbrKey(text), out var dimId, out var memberId)
                ? (dimId, memberId) : (0, 0);
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
                    if (SheetMetaParser.TryParseDimMbr(SheetMetaParser.ExtractDimMbrKey(text), out var dimId, out var memberId))
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


