using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialPersonal("MyOlap");

var path = @"c:\Users\nmali\OneDrive\Desktop\MyOlap v1.0\Prototype\MS-Excel-ERP-Analytics-main\MS-Excel-ERP-Analytics\Sales_Data Sample.xlsx";
using var pkg = new ExcelPackage(new FileInfo(path));

foreach (var ws in pkg.Workbook.Worksheets)
{
    Console.WriteLine($"=== {ws.Name} (rows: {ws.Dimension?.Rows}, cols: {ws.Dimension?.Columns}) ===");
    var cols = ws.Dimension?.Columns ?? 0;
    var rows = Math.Min(ws.Dimension?.Rows ?? 0, 8);
    for (int r = 1; r <= rows; r++)
    {
        var vals = new List<string>();
        for (int c = 1; c <= cols; c++)
            vals.Add(ws.Cells[r, c].Text ?? "");
        Console.WriteLine(string.Join(" | ", vals));
    }
    Console.WriteLine($"... total rows: {ws.Dimension?.Rows}");
    Console.WriteLine();
}