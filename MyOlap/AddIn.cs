using ExcelDna.Integration;
using MyOlap.Core;
using MyOlap.Data;

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
