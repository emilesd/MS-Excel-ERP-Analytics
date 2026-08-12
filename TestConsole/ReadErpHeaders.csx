// Quick script to inspect ERP file structure
// Run with: dotnet script ReadErpHeaders.csx (or compile/run via dotnet)
// Actually run via TestConsole manually
using OfficeOpenXml;
using System.IO;

ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
var path = @"C:\MyOlap\TestData\Sales_Data ERP.xlsx";
using var pkg = new ExcelPackage(new FileInfo(path));

Console.WriteLine("=== Sheets ===");
foreach (var ws2 in pkg.Workbook.Worksheets)
    Console.WriteLine($"  {ws2.Name}: {ws2.Dimension?.End.Row ?? 0} rows");

var ws = pkg.Workbook.Worksheets[0];
Console.WriteLine($"\n=== Sheet 1: {ws.Name} ===");
Console.WriteLine("Headers:");
for (int c = 1; c <= ws.Dimension.End.Column; c++)
    Console.WriteLine($"  [{c}] {ws.Cells[1,c].Text}");

Console.WriteLine("\nSample rows 2-4:");
for (int r = 2; r <= Math.Min(4, ws.Dimension.End.Row); r++)
{
    var vals = new List<string>();
    for (int c = 1; c <= ws.Dimension.End.Column; c++)
        vals.Add(ws.Cells[r,c].Text);
    Console.WriteLine($"  Row {r}: {string.Join(" | ", vals)}");
}

// Find unique values per column (first 5000 rows sample)
Console.WriteLine("\n=== Unique values per column (sample 5000 rows) ===");
for (int c = 1; c <= ws.Dimension.End.Column; c++)
{
    var unique = new HashSet<string>();
    for (int r = 2; r <= Math.Min(5001, ws.Dimension.End.Row); r++)
        unique.Add(ws.Cells[r,c].Text);
    var header = ws.Cells[1,c].Text;
    if (unique.Count <= 30)
        Console.WriteLine($"  {header}: [{string.Join(", ", unique.OrderBy(x => x))}]");
    else
        Console.WriteLine($"  {header}: {unique.Count} distinct values (too many to list)");
}
