using System.Globalization;
using System.IO;
using CsvHelper;
using CsvHelper.Configuration;
using MyOlap.Core;
using OfficeOpenXml;

namespace MyOlap.Data;

/// <summary>
/// Loads fact data from Excel, CSV, or TXT files into the SQLite model.
/// Supports column mapping and batch insertion for 100k-row performance.
/// </summary>
public class DataLoader
{
    private readonly SqliteRepository _repo = SqliteRepository.Instance;
    public int SkippedRows { get; private set; }

    /// <summary>
    /// Column mapping: maps each source-file column index to a dimension ID.
    /// The last mapped column is treated as the value/measure column.
    /// </summary>
    public class ColumnMapping
    {
        public Dictionary<int, long> ColumnToDimension { get; set; } = new();
        public int ValueColumnIndex { get; set; }
        public bool ValueIsText { get; set; }
    }

    /// <summary>
    /// Reads the header row from a file to allow the user to build a column mapping.
    /// </summary>
    public List<string> ReadHeaders(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" or ".xls" => ReadExcelHeaders(filePath),
            ".csv" => ReadCsvHeaders(filePath, ','),
            ".txt" => ReadCsvHeaders(filePath, '\t'),
            _ => throw new NotSupportedException($"Unsupported file format: {ext}")
        };
    }

    /// <summary>
    /// Returns the user-facing alert when one or more model dimensions are not mapped
    /// in the data source, or null when every dimension has a column.
    /// </summary>
    public static string? GetMissingDimensionAlert(IReadOnlyList<Dimension> dims, ColumnMapping mapping)
    {
        if (dims == null || dims.Count == 0 || mapping == null) return null;
        var mapped = new HashSet<long>(mapping.ColumnToDimension.Values);
        var missing = dims.Where(d => !mapped.Contains(d.Id)).Select(d => d.Name).ToList();
        if (missing.Count == 0) return null;
        return $"Data for dimension {string.Join(", ", missing)} not provided. Please provide data for all dimensions in the data source";
    }

    /// <summary>
    /// Loads all data rows from the file using the provided mapping.
    /// Member names are resolved to IDs; unknown members are skipped.
    /// Data is inserted in a single transaction for performance.
    /// Requires every model dimension to be mapped to a source column.
    /// </summary>
    public int LoadData(string filePath, long modelId, ColumnMapping mapping)
    {
        var dims = _repo.GetDimensions(modelId);
        var missingAlert = GetMissingDimensionAlert(dims, mapping);
        if (missingAlert != null)
            throw new InvalidOperationException(missingAlert);

        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var rows = ext switch
        {
            ".xlsx" or ".xls" => ReadExcelRows(filePath),
            ".csv" => ReadCsvRows(filePath, ','),
            ".txt" => ReadCsvRows(filePath, '\t'),
            _ => throw new NotSupportedException($"Unsupported file format: {ext}")
        };

        var dimOrder = dims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();

        var memberCache = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var dim in dims)
        {
            foreach (var m in _repo.GetMembers(dim.Id))
                memberCache[$"{dim.Id}|{m.Name}"] = m.Id;
        }

        var facts = new List<FactData>();
        int loaded = 0;
        int skipped = 0;

        foreach (var row in rows)
        {
            var memberIds = new Dictionary<long, long>();
            bool skip = false;

            foreach (var (colIdx, dimId) in mapping.ColumnToDimension)
            {
                if (colIdx >= row.Count) { skip = true; break; }
                var rawName = row[colIdx]?.Trim() ?? "";
                if (string.IsNullOrEmpty(rawName)) { skip = true; break; }

                var memberName = rawName;
                if (decimal.TryParse(rawName, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var numVal)
                    && numVal == Math.Truncate(numVal))
                    memberName = ((long)numVal).ToString();

                var cacheKey = $"{dimId}|{memberName}";
                if (!memberCache.TryGetValue(cacheKey, out var memberId))
                {
                    var altKey = $"{dimId}|{rawName}";
                    if (!memberCache.TryGetValue(altKey, out memberId))
                    {
                        skip = true;
                        skipped++;
                        break;
                    }
                }
                memberIds[dimId] = memberId;
            }

            if (skip) continue;

            // All dimensions were validated as mapped; every dimId must have a member.
            if (dimOrder.Any(dimId => !memberIds.ContainsKey(dimId)))
            {
                skipped++;
                continue;
            }

            var key = OlapEngine.BuildMemberKey(dimOrder, memberIds);
            var valueStr = mapping.ValueColumnIndex < row.Count
                ? row[mapping.ValueColumnIndex]?.Trim() ?? ""
                : "";

            var fact = new FactData { ModelId = modelId, MemberKey = key };
            if (mapping.ValueIsText)
            {
                fact.TextValue = valueStr;
                fact.DataType = MeasureDataType.Text;
            }
            else
            {
                fact.NumericValue = decimal.TryParse(valueStr, NumberStyles.Any,
                    CultureInfo.InvariantCulture, out var nv) ? nv : null;
                fact.DataType = MeasureDataType.Numeric;
            }
            facts.Add(fact);
            loaded++;
        }

        _repo.InsertFactBatch(modelId, facts);
        if (skipped > 0)
            SkippedRows = skipped;
        return loaded;
    }

    #region File readers

    private static List<string> ReadExcelHeaders(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
        using var pkg = new ExcelPackage(new FileInfo(filePath));
        var ws = pkg.Workbook.Worksheets[0];
        var headers = new List<string>();
        for (int col = 1; col <= ws.Dimension.End.Column; col++)
            headers.Add(ws.Cells[1, col].Text.Trim());
        return headers;
    }

    private static List<List<string>> ReadExcelRows(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("MyOlap");
        using var pkg = new ExcelPackage(new FileInfo(filePath));
        var rows = new List<List<string>>();
        var firstSheet = pkg.Workbook.Worksheets[0];
        int colCount = firstSheet.Dimension.End.Column;
        var baseHeaders = new List<string>();
        for (int c = 1; c <= colCount; c++)
            baseHeaders.Add(firstSheet.Cells[1, c].Text.Trim().ToLowerInvariant());

        foreach (var ws in pkg.Workbook.Worksheets)
        {
            if (ws.Dimension == null) continue;
            bool sameHeaders = ws.Dimension.End.Column == colCount;
            if (sameHeaders)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    if (!ws.Cells[1, c].Text.Trim().Equals(baseHeaders[c - 1], StringComparison.OrdinalIgnoreCase))
                    { sameHeaders = false; break; }
                }
            }
            if (!sameHeaders) continue;

            for (int r = 2; r <= ws.Dimension.End.Row; r++)
            {
                var row = new List<string>();
                for (int c = 1; c <= ws.Dimension.End.Column; c++)
                    row.Add(ws.Cells[r, c].Text);
                rows.Add(row);
            }
        }
        return rows;
    }

    private static List<string> ReadCsvHeaders(string filePath, char delimiter)
    {
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            HasHeaderRecord = true
        };
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, config);
        csv.Read();
        csv.ReadHeader();
        return csv.HeaderRecord?.ToList() ?? new List<string>();
    }

    private static List<List<string>> ReadCsvRows(string filePath, char delimiter)
    {
        var config = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            HasHeaderRecord = true
        };
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, config);
        csv.Read();
        csv.ReadHeader();
        var rows = new List<List<string>>();
        while (csv.Read())
        {
            var row = new List<string>();
            for (int i = 0; i < csv.Parser.Count; i++)
                row.Add(csv.GetField(i) ?? "");
            rows.Add(row);
        }
        return rows;
    }

    #endregion
}
