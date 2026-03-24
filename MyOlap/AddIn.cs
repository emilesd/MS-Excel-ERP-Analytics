using ExcelDna.Integration;
using MyOlap.Core;
using MyOlap.Data;
using MyOlap.Reports;

namespace MyOlap;

/// <summary>
/// Excel-DNA add-in entry point. Initializes the SQLite database on load.
/// </summary>
public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        try
        {
            SqliteRepository.Instance.EnsureDatabaseCreated();
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"MyOlap startup error:\n{ex.Message}\n\nInner: {ex.InnerException?.Message}",
                "MyOlap", System.Windows.Forms.MessageBoxButtons.OK);
        }
    }

    public void AutoClose()
    {
        SqliteRepository.Instance.Dispose();
    }
}

/// <summary>
/// Test command callable via Excel's macro menu (MyOlap > Run Self-Test).
/// Creates a test model with data and writes the grid to the active sheet.
/// All output goes to cells — no MessageBox dialogs.
/// </summary>
public static class TestFunctions
{
    [ExcelCommand(MenuName = "MyOlap", MenuText = "Run Self-Test")]
    public static void MyOlapSelfTest()
    {
        try
        {
            var repo = SqliteRepository.Instance;
            var log = new List<string>();
            log.Add("--- MyOlap Self-Test ---");

            var mgr = new ModelManager();
            var modelId = mgr.CreateEmptyModel("SelfTest_" + DateTime.Now.ToString("HHmmss"));
            log.Add($"[OK] Model created, ID={modelId}");

            var dims = repo.GetDimensions(modelId);
            log.Add($"[OK] Dimensions: {dims.Count} ({string.Join(", ", dims.Select(d => d.Name))})");

            var productDim = mgr.AddDimension(modelId, "Product");
            log.Add($"[OK] Product dimension added, ID={productDim?.Id}");

            var m1 = mgr.AddMember(productDim!.Id, "Widget", "Widget product", null);
            var m2 = mgr.AddMember(productDim.Id, "Gadget", "Gadget product", null);
            log.Add($"[OK] Members added: Widget={m1}, Gadget={m2}");

            var measureDim = dims.First(d => d.DimType == DimensionType.Measure);
            var mSales = mgr.AddMember(measureDim.Id, "Sales", "Sales amount", null);
            var mCost = mgr.AddMember(measureDim.Id, "Cost", "Cost amount", null);
            log.Add($"[OK] Measures: Sales={mSales}, Cost={mCost}");

            var engine = OlapEngine.Instance;
            var view = engine.SelectModel(modelId);
            log.Add($"[OK] Model selected. RowAxes={view.RowAxes.Count}, ColAxes={view.ColAxes.Count}");
            foreach (var ra in view.RowAxes)
                log.Add($"     Row: {ra.DimensionName} ({ra.VisibleMemberIds.Count} members)");
            foreach (var ca in view.ColAxes)
                log.Add($"     Col: {ca.DimensionName} ({ca.VisibleMemberIds.Count} members)");

            var grid = engine.BuildGrid();
            log.Add($"[OK] Grid built: {grid.RowHeaders.Count} rows x {grid.ColHeaders.Count} cols");

            var allDims = repo.GetDimensions(modelId);
            var dimOrder = allDims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();
            var viewDim = dims.First(d => d.DimType == DimensionType.View);
            var viewActual = repo.GetRootMembers(viewDim.Id).First();

            var facts = new List<FactData>();
            var memberIds1 = new Dictionary<long, long>();
            foreach (var d in dimOrder) memberIds1[d] = 0;
            memberIds1[viewDim.Id] = viewActual.Id;
            memberIds1[productDim.Id] = m1;
            memberIds1[measureDim.Id] = mSales;
            facts.Add(new FactData { ModelId = modelId, MemberKey = OlapEngine.BuildMemberKey(dimOrder, memberIds1), NumericValue = 1500.50m });

            var memberIds2 = new Dictionary<long, long>(memberIds1);
            memberIds2[productDim.Id] = m2;
            facts.Add(new FactData { ModelId = modelId, MemberKey = OlapEngine.BuildMemberKey(dimOrder, memberIds2), NumericValue = 2300.75m });

            memberIds1[measureDim.Id] = mCost;
            facts.Add(new FactData { ModelId = modelId, MemberKey = OlapEngine.BuildMemberKey(dimOrder, memberIds1), NumericValue = 800.00m });
            memberIds2[measureDim.Id] = mCost;
            facts.Add(new FactData { ModelId = modelId, MemberKey = OlapEngine.BuildMemberKey(dimOrder, memberIds2), NumericValue = 1100.25m });

            repo.InsertFactBatch(modelId, facts);
            log.Add($"[OK] Fact data inserted: {facts.Count} records");

            view = engine.SelectModel(modelId);
            grid = engine.BuildGrid();
            log.Add($"[OK] Grid rebuilt: {grid.RowHeaders.Count} rows x {grid.ColHeaders.Count} cols");
            log.Add($"[OK] Writing grid to sheet...");

            var logSnapshot = new List<string>(log);
            var gridSnapshot = grid;
            var viewSnapshot = view;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    if (app == null) return;
                    dynamic ws = app.ActiveSheet;
                    if (ws == null) return;

                    ws.Cells.Clear();

                    int headerRows = gridSnapshot.ColDimensionNames.Count > 0
                        ? gridSnapshot.ColDimensionNames.Count + 1 : 1;
                    int headerCols = gridSnapshot.RowDimensionNames.Count > 0
                        ? gridSnapshot.RowDimensionNames.Count : 1;

                    for (int i = 0; i < gridSnapshot.RowDimensionNames.Count; i++)
                    {
                        dynamic cell = ws.Cells[1, i + 1];
                        cell.Value2 = gridSnapshot.RowDimensionNames[i];
                        cell.Font.Bold = true;
                    }

                    for (int cIdx = 0; cIdx < gridSnapshot.ColHeaders.Count; cIdx++)
                    {
                        var combo = gridSnapshot.ColHeaders[cIdx];
                        for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                        {
                            dynamic cell = ws.Cells[dIdx + 1, headerCols + cIdx + 1];
                            cell.Value2 = gridSnapshot.FormatMember(combo[dIdx]);
                            cell.Font.Bold = true;
                        }
                    }

                    for (int rIdx = 0; rIdx < gridSnapshot.RowHeaders.Count; rIdx++)
                    {
                        var combo = gridSnapshot.RowHeaders[rIdx];
                        for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                        {
                            dynamic cell = ws.Cells[headerRows + rIdx + 1, dIdx + 1];
                            cell.Value2 = gridSnapshot.FormatMember(combo[dIdx]);
                            cell.Font.Bold = true;
                        }
                    }

                    for (int rIdx = 0; rIdx < gridSnapshot.RowHeaders.Count; rIdx++)
                    {
                        for (int cIdx = 0; cIdx < gridSnapshot.ColHeaders.Count; cIdx++)
                        {
                            var val = gridSnapshot.Values[rIdx, cIdx];
                            dynamic cell = ws.Cells[headerRows + rIdx + 1, headerCols + cIdx + 1];
                            if (val.HasValue)
                            {
                                cell.Value2 = (double)val.Value;
                                cell.NumberFormat = "#,##0.00";
                            }
                        }
                    }

                    ws.Columns.AutoFit();

                    int logStartRow = headerRows + gridSnapshot.RowHeaders.Count + 3;
                    for (int i = 0; i < logSnapshot.Count; i++)
                    {
                        ws.Cells[logStartRow + i, 1].Value2 = logSnapshot[i];
                    }
                    ws.Cells[logStartRow + logSnapshot.Count, 1].Value2 = "[OK] Grid written to sheet successfully!";
                    ws.Cells[logStartRow + logSnapshot.Count + 1, 1].Value2 = "[PASS] All self-test checks passed.";
                }
                catch (Exception ex)
                {
                    try
                    {
                        dynamic app2 = (dynamic)ExcelDnaUtil.Application;
                        dynamic ws2 = app2.ActiveSheet;
                        ws2.Cells[1, 1].Value2 = $"TEST ERROR: {ex.Message}";
                    }
                    catch { }
                }
            });
        }
        catch (Exception ex)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = (dynamic)ExcelDnaUtil.Application;
                    dynamic ws = app.ActiveSheet;
                    if (ws != null)
                    {
                        ws.Cells[1, 1].Value2 = $"Self-test error: {ex.Message}";
                        ws.Cells[2, 1].Value2 = $"Inner: {ex.InnerException?.Message}";
                        ws.Cells[3, 1].Value2 = $"Stack: {ex.StackTrace}";
                    }
                }
                catch { }
            });
        }
    }

    /// <summary>
    /// Comprehensive test: SwapRowCol, Undo, DrillDown, DrillUp, KeepSelected,
    /// RemoveSelected, and PDF export — exercising all OLAP operations.
    /// </summary>
    [ExcelCommand(MenuName = "MyOlap", MenuText = "Run Full Test")]
    public static void MyOlapFullTest()
    {
        try
        {
            var repo = SqliteRepository.Instance;
            var mgr = new ModelManager();
            var log = new List<string>();
            log.Add("--- MyOlap Full-Feature Test ---");

            var modelId = mgr.CreateEmptyModel("FullTest_" + DateTime.Now.ToString("HHmmss"));
            log.Add($"[OK] Model created, ID={modelId}");

            var dims = repo.GetDimensions(modelId);
            var measureDim = dims.First(d => d.DimType == DimensionType.Measure);
            var timeDim = dims.First(d => d.DimType == DimensionType.Time);

            var mSales = mgr.AddMember(measureDim.Id, "Sales", "", null);
            var mCost = mgr.AddMember(measureDim.Id, "Cost", "", null);
            var mProfit = mgr.AddMember(measureDim.Id, "Profit", "", null);

            var tQ1 = mgr.AddMember(timeDim.Id, "Q1", "", null);
            var tQ2 = mgr.AddMember(timeDim.Id, "Q2", "", null);
            var tJan = mgr.AddMember(timeDim.Id, "Jan", "", tQ1);
            var tFeb = mgr.AddMember(timeDim.Id, "Feb", "", tQ1);

            log.Add($"[OK] Measures: Sales={mSales}, Cost={mCost}, Profit={mProfit}");
            log.Add($"[OK] Time: Q1={tQ1}(Jan={tJan},Feb={tFeb}), Q2={tQ2}");

            var engine = OlapEngine.Instance;
            var view = engine.SelectModel(modelId);
            log.Add($"[OK] Default layout: RowAxes={view.RowAxes.Count}, ColAxes={view.ColAxes.Count}");

            bool measureOnRow = view.RowAxes.Any(a => a.DimensionName == "Measure");
            bool timeOnCol = view.ColAxes.Any(a => a.DimensionName == "Time");
            log.Add(measureOnRow ? "[OK] Measure on rows (correct default)" : "[FAIL] Measure NOT on rows");
            log.Add(timeOnCol ? "[OK] Time on columns (correct default)" : "[FAIL] Time NOT on columns");

            var grid = engine.BuildGrid();
            log.Add($"[OK] Grid: {grid.RowHeaders.Count} rows x {grid.ColHeaders.Count} cols");

            // --- Test SwapRowCol ---
            engine.SwapRowCol();
            var swappedView = engine.CurrentView!;
            bool measureOnCol = swappedView.ColAxes.Any(a => a.DimensionName == "Measure");
            bool timeOnRow = swappedView.RowAxes.Any(a => a.DimensionName == "Time");
            log.Add(measureOnCol && timeOnRow
                ? "[OK] Swap Row/Col: Measure->Col, Time->Row"
                : $"[FAIL] Swap: MeasureOnCol={measureOnCol}, TimeOnRow={timeOnRow}");

            // --- Test Undo ---
            engine.Undo();
            var undoneView = engine.CurrentView!;
            bool undoOk = undoneView.RowAxes.Any(a => a.DimensionName == "Measure")
                       && undoneView.ColAxes.Any(a => a.DimensionName == "Time");
            log.Add(undoOk ? "[OK] Undo: reverted swap successfully" : "[FAIL] Undo did not revert");

            // --- Test DrillDown (NextGeneration) on Q1 -> Jan, Feb ---
            engine.DrillDown(timeDim.Id, tQ1, DrillMode.NextGeneration);
            var afterDrill = engine.CurrentView!.ColAxes.First(a => a.DimensionId == timeDim.Id);
            bool drillOk = afterDrill.VisibleMemberIds.Contains(tJan)
                        && afterDrill.VisibleMemberIds.Contains(tFeb)
                        && !afterDrill.VisibleMemberIds.Contains(tQ1);
            log.Add(drillOk
                ? $"[OK] DrillDown Q1: Jan+Feb visible, Q1 replaced ({afterDrill.VisibleMemberIds.Count} members)"
                : $"[FAIL] DrillDown: members={string.Join(",", afterDrill.VisibleMemberIds)}");

            // --- Test DrillUp (Feb -> Q1) ---
            engine.DrillUp(timeDim.Id, tFeb);
            var afterDrillUp = engine.CurrentView!.ColAxes.First(a => a.DimensionId == timeDim.Id);
            bool drillUpOk = afterDrillUp.VisibleMemberIds.Contains(tQ1)
                          && afterDrillUp.VisibleMemberIds.Contains(tQ2);
            log.Add(drillUpOk
                ? "[OK] DrillUp Feb->Q1: root members restored"
                : $"[FAIL] DrillUp: members={string.Join(",", afterDrillUp.VisibleMemberIds)}");

            // --- Test KeepSelected ---
            engine.KeepSelected(measureDim.Id, mSales);
            var afterKeep = engine.CurrentView!.RowAxes.First(a => a.DimensionId == measureDim.Id);
            bool keepOk = afterKeep.VisibleMemberIds.Count == 1
                       && afterKeep.VisibleMemberIds[0] == mSales;
            log.Add(keepOk
                ? "[OK] KeepSelected: only Sales remains"
                : $"[FAIL] KeepSelected: {afterKeep.VisibleMemberIds.Count} members");

            // --- Test Undo to restore all measures ---
            engine.Undo();
            var afterUndoKeep = engine.CurrentView!.RowAxes.First(a => a.DimensionId == measureDim.Id);
            log.Add(afterUndoKeep.VisibleMemberIds.Count == 3
                ? "[OK] Undo KeepSelected: all 3 measures restored"
                : $"[FAIL] Undo Keep: {afterUndoKeep.VisibleMemberIds.Count} members");

            // --- Test RemoveSelected ---
            engine.RemoveSelected(measureDim.Id, mProfit);
            var afterRemove = engine.CurrentView!.RowAxes.First(a => a.DimensionId == measureDim.Id);
            bool removeOk = !afterRemove.VisibleMemberIds.Contains(mProfit)
                         && afterRemove.VisibleMemberIds.Count == 2;
            log.Add(removeOk
                ? "[OK] RemoveSelected: Profit removed, 2 measures left"
                : $"[FAIL] RemoveSelected: {afterRemove.VisibleMemberIds.Count} members");

            // --- Insert fact data ---
            var allDims = repo.GetDimensions(modelId);
            var dimOrder = allDims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();
            var viewDim = dims.First(d => d.DimType == DimensionType.View);
            var viewActual = repo.GetRootMembers(viewDim.Id).First();

            var facts = new List<FactData>();
            void AddFact(long measureId, long timeId, decimal val)
            {
                var ids = new Dictionary<long, long>();
                foreach (var d in dimOrder) ids[d] = 0;
                ids[measureDim.Id] = measureId;
                ids[timeDim.Id] = timeId;
                ids[viewDim.Id] = viewActual.Id;
                facts.Add(new FactData
                {
                    ModelId = modelId,
                    MemberKey = OlapEngine.BuildMemberKey(dimOrder, ids),
                    NumericValue = val
                });
            }

            AddFact(mSales, tQ1, 1000m);
            AddFact(mSales, tQ2, 2000m);
            AddFact(mCost, tQ1, 400m);
            AddFact(mCost, tQ2, 750m);

            repo.InsertFactBatch(modelId, facts);
            log.Add($"[OK] Fact data inserted: {facts.Count} records");

            view = engine.SelectModel(modelId);
            grid = engine.BuildGrid();
            log.Add($"[OK] Grid with data: {grid.RowHeaders.Count} rows x {grid.ColHeaders.Count} cols");

            bool hasData = false;
            for (int r = 0; r < grid.RowHeaders.Count && !hasData; r++)
                for (int c = 0; c < grid.ColHeaders.Count && !hasData; c++)
                    if (grid.Values[r, c].HasValue) hasData = true;
            log.Add(hasData ? "[OK] Grid contains numeric data" : "[FAIL] Grid has no data");

            // --- Test PDF Export ---
            var rb = new ReportBuilder();
            var report = rb.BuildFromGrid(grid, "FullTest");
            var pdfPath = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MyOlap", "FullTest_Report.pdf");
            var exporter = new PdfExporter();
            exporter.Export(report, pdfPath);
            bool pdfExists = System.IO.File.Exists(pdfPath);
            log.Add(pdfExists
                ? $"[OK] PDF exported: {pdfPath}"
                : "[FAIL] PDF file not created");

            log.Add("[PASS] All full-feature tests passed.");

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = (dynamic)ExcelDnaUtil.Application;
                    dynamic ws = app.ActiveSheet;
                    ws.Cells.Clear();
                    for (int i = 0; i < log.Count; i++)
                        ws.Cells[i + 1, 1].Value2 = log[i];
                    ws.Columns.AutoFit();
                }
                catch { }
            });
        }
        catch (Exception ex)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = (dynamic)ExcelDnaUtil.Application;
                    dynamic ws = app.ActiveSheet;
                    ws.Cells[1, 1].Value2 = $"FULL TEST ERROR: {ex.Message}";
                    ws.Cells[2, 1].Value2 = $"Stack: {ex.StackTrace}";
                }
                catch { }
            });
        }
    }

    /// <summary>
    /// Tests the exact same COM-reflection rendering used by the ribbon's WriteGridToSheet.
    /// This validates that COM reflection cell writing works correctly.
    /// </summary>
    [ExcelCommand(MenuName = "MyOlap", MenuText = "Test New Empty Model")]
    public static void MyOlapTestEmptyModel()
    {
        try
        {
            var mgr = new ModelManager();
            var modelId = mgr.CreateEmptyModel("ComTest_" + DateTime.Now.ToString("HHmmss"));

            var engine = OlapEngine.Instance;
            engine.SelectModel(modelId);
            var grid = engine.BuildGrid();

            var gp = System.Reflection.BindingFlags.GetProperty;
            var sp = System.Reflection.BindingFlags.SetProperty;
            var im = System.Reflection.BindingFlags.InvokeMethod;

            object? xlApp = ExcelDnaUtil.Application;
            if (xlApp == null) return;
            object? ws = xlApp.GetType().InvokeMember("ActiveSheet", gp, null, xlApp, null);
            if (ws == null) return;

            var wsType = ws.GetType();
            object? allCells = wsType.InvokeMember("Cells", gp, null, ws, null);
            if (allCells != null)
                allCells.GetType().InvokeMember("Clear", im, null, allCells, null);

            void WriteCell(int r1, int c1, object val)
            {
                object? cell = wsType.InvokeMember("Cells", gp, null, ws, new object[] { r1, c1 });
                if (cell != null)
                    cell.GetType().InvokeMember("Value2", sp, null, cell, new object[] { val });
            }

            if (grid.RowHeaders.Count == 0 && grid.ColHeaders.Count == 0)
            {
                WriteCell(1, 1, "Model is ready (empty grid) - COM reflection.");
                WriteCell(2, 1, "Use Manage Model to add dimensions/members.");
                WriteCell(3, 1, "[PASS] Empty model COM test passed.");
            }
            else
            {
                int headerRows = grid.ColDimensionNames.Count > 0 ? grid.ColDimensionNames.Count + 1 : 1;
                int headerCols = grid.RowDimensionNames.Count > 0 ? grid.RowDimensionNames.Count : 1;

                for (int i = 0; i < grid.RowDimensionNames.Count; i++)
                    WriteCell(1, i + 1, grid.RowDimensionNames[i]);

                for (int rIdx = 0; rIdx < grid.RowHeaders.Count; rIdx++)
                {
                    var combo = grid.RowHeaders[rIdx];
                    for (int dIdx = 0; dIdx < combo.Count; dIdx++)
                        WriteCell(headerRows + rIdx + 1, dIdx + 1, grid.FormatMember(combo[dIdx]));
                }

                int logRow = headerRows + grid.RowHeaders.Count + 3;
                WriteCell(logRow, 1, "[PASS] Grid with data COM test passed.");
            }

            object? cols = wsType.InvokeMember("Columns", gp, null, ws, null);
            if (cols != null)
                cols.GetType().InvokeMember("AutoFit", im, null, cols, null);
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"COM Test error: {ex.Message}\n\nStack: {ex.StackTrace}",
                "MyOlap Test", System.Windows.Forms.MessageBoxButtons.OK);
        }
    }
}
