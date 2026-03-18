using MyOlap.Core;
using MyOlap.Data;

namespace MyOlap.Reports;

/// <summary>
/// Prepares structured report data from the current OLAP view.
/// </summary>
public class ReportBuilder
{
    public ReportData BuildFromGrid(GridResult grid, string modelName)
    {
        var report = new ReportData
        {
            Title = $"MyOlap Report – {modelName}",
            GeneratedUtc = DateTime.UtcNow,
            RowDimensionNames = grid.RowDimensionNames,
            ColDimensionNames = grid.ColDimensionNames
        };

        foreach (var combo in grid.ColHeaders)
            report.ColumnLabels.Add(string.Join(" / ", combo.Select(m => grid.FormatMember(m))));

        foreach (var combo in grid.RowHeaders)
            report.RowLabels.Add(string.Join(" / ", combo.Select(m => grid.FormatMember(m))));

        report.Values = new string[grid.RowHeaders.Count, grid.ColHeaders.Count];
        for (int r = 0; r < grid.RowHeaders.Count; r++)
        {
            for (int c = 0; c < grid.ColHeaders.Count; c++)
            {
                report.Values[r, c] = grid.Values[r, c]?.ToString("N2") ?? "";
            }
        }

        return report;
    }
}

public class ReportData
{
    public string Title { get; set; } = "";
    public DateTime GeneratedUtc { get; set; }
    public List<string> RowDimensionNames { get; set; } = new();
    public List<string> ColDimensionNames { get; set; } = new();
    public List<string> ColumnLabels { get; set; } = new();
    public List<string> RowLabels { get; set; } = new();
    public string[,] Values { get; set; } = new string[0, 0];
}
