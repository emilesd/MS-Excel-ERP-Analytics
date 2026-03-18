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
    /// Loads all data rows from the file using the provided mapping.
    /// Member names are resolved to IDs; unknown members are auto-created.
    /// Data is inserted in a single transaction for performance.
    /// </summary>
    public int LoadData(string filePath, long modelId, ColumnMapping mapping)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var rows = ext switch
        {
            ".xlsx" or ".xls" => ReadExcelRows(filePath),
            ".csv" => ReadCsvRows(filePath, ','),
            ".txt" => ReadCsvRows(filePath, '\t'),
            _ => throw new NotSupportedException($"Unsupported file format: {ext}")
        };

        var dims = _repo.GetDimensions(modelId);
        var dimOrder = dims.OrderBy(d => d.SortOrder).Select(d => d.Id).ToList();

        var memberCache = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var dim in dims)
        {
            foreach (var m in _repo.GetMembers(dim.Id))
                memberCache[$"{dim.Id}|{m.Name}"] = m.Id;
        }

        var facts = new List<FactData>();
        int loaded = 0;

        foreach (var row in rows)
        {
            var memberIds = new Dictionary<long, long>();
            bool skip = false;

            foreach (var (colIdx, dimId) in mapping.ColumnToDimension)
            {
                if (colIdx >= row.Count) { skip = true; break; }
                var memberName = row[colIdx]?.Trim() ?? "";
                if (string.IsNullOrEmpty(memberName)) { skip = true; break; }

                var cacheKey = $"{dimId}|{memberName}";
                if (!memberCache.TryGetValue(cacheKey, out var memberId))
                {
                    memberId = _repo.InsertMember(new Member
                    {
                        DimensionId = dimId,
                        Name = memberName,
                        Description = memberName,
                        Level = 0,
                        SortOrder = memberCache.Count(kv => kv.Key.StartsWith($"{dimId}|"))
                    });
                    memberCache[cacheKey] = memberId;
                }
                memberIds[dimId] = memberId;
            }

            if (skip) continue;

            foreach (var dimId in dimOrder)
            {
                if (!memberIds.ContainsKey(dimId))
                    memberIds[dimId] = 0;
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
        var ws = pkg.Workbook.Worksheets[0];
        var rows = new List<List<string>>();
        for (int r = 2; r <= ws.Dimension.End.Row; r++)
        {
            var row = new List<string>();
            for (int c = 1; c <= ws.Dimension.End.Column; c++)
                row.Add(ws.Cells[r, c].Text);
            rows.Add(row);
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
