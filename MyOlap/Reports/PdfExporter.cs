using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace MyOlap.Reports;

/// <summary>
/// Renders a ReportData object to a PDF file using PdfSharpCore.
/// </summary>
public class PdfExporter
{
    private const double MarginLeft = 40;
    private const double MarginTop = 40;
    private const double CellPadding = 4;
    private const double RowHeight = 18;
    private const double ColWidth = 100;
    private const double RowHeaderWidth = 150;

    public void Export(ReportData report, string outputPath)
    {
        var doc = new PdfDocument();
        doc.Info.Title = report.Title;

        var page = doc.AddPage();
        page.Orientation = PdfSharpCore.PageOrientation.Landscape;
        var gfx = XGraphics.FromPdfPage(page);

        var titleFont = new XFont("Arial", 14, XFontStyle.Bold);
        var headerFont = new XFont("Arial", 9, XFontStyle.Bold);
        var cellFont = new XFont("Arial", 9, XFontStyle.Regular);

        double y = MarginTop;

        gfx.DrawString(report.Title, titleFont, XBrushes.Black,
            new XPoint(MarginLeft, y));
        y += 22;

        gfx.DrawString($"Generated: {report.GeneratedUtc:yyyy-MM-dd HH:mm} UTC",
            cellFont, XBrushes.DarkGray, new XPoint(MarginLeft, y));
        y += 24;

        double x = MarginLeft + RowHeaderWidth;
        foreach (var colLabel in report.ColumnLabels)
        {
            gfx.DrawString(colLabel, headerFont, XBrushes.Black,
                new XRect(x, y, ColWidth, RowHeight), XStringFormats.CenterLeft);
            x += ColWidth;
        }
        y += RowHeight + 2;

        var lineY = y;
        gfx.DrawLine(XPens.Gray, MarginLeft, lineY, x, lineY);
        y += 4;

        for (int r = 0; r < report.RowLabels.Count; r++)
        {
            if (y + RowHeight > page.Height - MarginTop)
            {
                page = doc.AddPage();
                page.Orientation = PdfSharpCore.PageOrientation.Landscape;
                gfx = XGraphics.FromPdfPage(page);
                y = MarginTop;
            }

            gfx.DrawString(report.RowLabels[r], headerFont, XBrushes.Black,
                new XRect(MarginLeft, y, RowHeaderWidth, RowHeight), XStringFormats.CenterLeft);

            x = MarginLeft + RowHeaderWidth;
            for (int c = 0; c < report.ColumnLabels.Count; c++)
            {
                var val = report.Values[r, c];
                gfx.DrawString(val, cellFont, XBrushes.Black,
                    new XRect(x, y, ColWidth, RowHeight), XStringFormats.CenterRight);
                x += ColWidth;
            }
            y += RowHeight;
        }

        doc.Save(outputPath);
    }
}
