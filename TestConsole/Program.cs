using System;
using System.IO;
using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
using var pkg = new ExcelPackage(new FileInfo(@"C:\MyOlap\Sales_Dimensions.xlsx"));
foreach (var ws in pkg.Workbook.Worksheets)
{
    Console.WriteLine($"Sheet: {ws.Name}  Rows: {ws.Dimension?.End.Row}  Cols: {ws.Dimension?.End.Column}");
    if (ws.Dimension != null)
    {
        for (int c = 1; c <= ws.Dimension.End.Column; c++)
            Console.Write($"  [{ws.Cells[1,c].Text}]");
        Console.WriteLine();
        for (int r = 2; r <= Math.Min(8, ws.Dimension.End.Row); r++)
        {
            for (int c = 1; c <= ws.Dimension.End.Column; c++)
                Console.Write($"  [{ws.Cells[r,c].Text}]");
            Console.WriteLine();
        }
        Console.WriteLine($"  ... ({ws.Dimension.End.Row - 1} data rows total)");
    }
    Console.WriteLine();
}