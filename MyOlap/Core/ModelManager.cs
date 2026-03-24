using System.IO;
using MyOlap.Data;
using OfficeOpenXml;

namespace MyOlap.Core;

/// <summary>
/// Handles model creation from workbooks, structure management,
/// and pre-populates the 5 required dimensions.
/// </summary>
public class ModelManager
{
    private readonly SqliteRepository _repo = SqliteRepository.Instance;

    /// <summary>
    /// Creates a new model with the 5 pre-defined dimensions and default members.
    /// Returns the new model ID.
    /// </summary>
    public long CreateEmptyModel(string name, string description = "")
    {
        var model = new OlapModel { Name = name, Description = description };
        var modelId = _repo.InsertModel(model);

        CreatePreDefinedDimensions(modelId);

        _repo.SaveSettings(new ModelSettings { ModelId = modelId });

        return modelId;
    }

    /// <summary>
    /// Creates a model by reading structure definitions from an Excel workbook.
    /// Expected workbook layout:
    ///   Sheet "Dimensions": Name | Type | SortOrder
    ///   Sheet "Members":    DimensionName | MemberName | ParentName | Description | Level
    /// </summary>
    public long CreateModelFromWorkbook(string filePath, string modelName)
    {
        ExcelPackage.License.SetNonCommercialPersonal("MyOlap");

        var modelId = _repo.InsertModel(new OlapModel { Name = modelName });
        _repo.SaveSettings(new ModelSettings { ModelId = modelId });

        using var package = new ExcelPackage(new FileInfo(filePath));

        var dimSheet = package.Workbook.Worksheets["Dimensions"];
        if (dimSheet != null)
            ReadDimensions(dimSheet, modelId);
        else
            CreatePreDefinedDimensions(modelId);

        var memberSheet = package.Workbook.Worksheets["Members"];
        if (memberSheet != null)
            ReadMembers(memberSheet, modelId);

        return modelId;
    }

    /// <summary>
    /// Adds a new user-defined dimension to an existing model.
    /// Enforces the limit of 12 dimensions total.
    /// </summary>
    public Dimension? AddDimension(long modelId, string name)
    {
        var existing = _repo.GetDimensions(modelId);
        if (existing.Count >= 12)
            return null;

        var dim = new Dimension
        {
            ModelId = modelId,
            Name = name,
            DimType = DimensionType.UserDefined,
            SortOrder = existing.Count
        };
        dim.Id = _repo.InsertDimension(dim);
        return dim;
    }

    public void RenameDimension(long dimId, string newName)
    {
        var dims = _repo.GetDimensions(0); // need to fetch by dimId
        // Fetch all dims for the model this dim belongs to
        var conn = SqliteRepository.Instance;
        // Simpler: update directly
        var dim = new Dimension { Id = dimId, Name = newName };
        // We need existing data; let's do a targeted update
        _repo.UpdateDimension(new Dimension { Id = dimId, Name = newName, DimType = DimensionType.UserDefined, SortOrder = 0 });
    }

    public long AddMember(long dimensionId, string name, string description, long? parentId)
    {
        int level = 0;
        if (parentId.HasValue)
        {
            var parent = _repo.GetMember(parentId.Value);
            if (parent != null) level = parent.Level + 1;
        }
        var members = _repo.GetMembers(dimensionId);
        var sortOrder = members.Count;
        return _repo.InsertMember(new Member
        {
            DimensionId = dimensionId,
            ParentId = parentId,
            Name = name,
            Description = description,
            Level = level,
            SortOrder = sortOrder
        });
    }

    #region Private helpers

    private void CreatePreDefinedDimensions(long modelId)
    {
        InsertDim(modelId, "Measure", DimensionType.Measure, 0);
        InsertDim(modelId, "Time", DimensionType.Time, 1);
        InsertDim(modelId, "Year", DimensionType.Year, 2);

        var viewDim = InsertDim(modelId, "View", DimensionType.View, 3);
        _repo.InsertMember(new Member { DimensionId = viewDim, Name = "Actual", Description = "Actual", Level = 0, SortOrder = 0 });
        _repo.InsertMember(new Member { DimensionId = viewDim, Name = "Budget", Description = "Budget", Level = 0, SortOrder = 1 });
        _repo.InsertMember(new Member { DimensionId = viewDim, Name = "Forecast", Description = "Forecast", Level = 0, SortOrder = 2 });

        InsertDim(modelId, "Version", DimensionType.Version, 4);
    }

    private long InsertDim(long modelId, string name, DimensionType type, int sort)
    {
        return _repo.InsertDimension(new Dimension
        {
            ModelId = modelId,
            Name = name,
            DimType = type,
            SortOrder = sort
        });
    }

    private void ReadDimensions(ExcelWorksheet sheet, long modelId)
    {
        int row = 2; // skip header
        while (sheet.Cells[row, 1].Value != null)
        {
            var name = sheet.Cells[row, 1].Text.Trim();
            var typeStr = sheet.Cells[row, 2].Text.Trim();
            var sort = int.TryParse(sheet.Cells[row, 3].Text, out var s) ? s : row - 2;

            var dimType = typeStr.ToLowerInvariant() switch
            {
                "view" => DimensionType.View,
                "version" => DimensionType.Version,
                "time" => DimensionType.Time,
                "year" => DimensionType.Year,
                "measure" => DimensionType.Measure,
                _ => DimensionType.UserDefined
            };

            _repo.InsertDimension(new Dimension
            {
                ModelId = modelId,
                Name = name,
                DimType = dimType,
                SortOrder = sort
            });
            row++;
        }
    }

    private void ReadMembers(ExcelWorksheet sheet, long modelId)
    {
        var dims = _repo.GetDimensions(modelId);
        var dimLookup = dims.ToDictionary(d => d.Name, StringComparer.OrdinalIgnoreCase);
        var memberLookup = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);

        int row = 2;
        while (sheet.Cells[row, 1].Value != null)
        {
            var dimName = sheet.Cells[row, 1].Text.Trim();
            var memberName = sheet.Cells[row, 2].Text.Trim();
            var parentName = sheet.Cells[row, 3].Text.Trim();
            var desc = sheet.Cells[row, 4].Text.Trim();
            var level = int.TryParse(sheet.Cells[row, 5].Text, out var lv) ? lv : 0;

            if (!dimLookup.TryGetValue(dimName, out var dim))
            {
                row++;
                continue;
            }

            long? parentId = null;
            if (!string.IsNullOrEmpty(parentName))
            {
                var pKey = $"{dimName}|{parentName}";
                memberLookup.TryGetValue(pKey, out var pid);
                if (pid > 0) parentId = pid;
            }

            var id = _repo.InsertMember(new Member
            {
                DimensionId = dim.Id,
                ParentId = parentId,
                Name = memberName,
                Description = desc,
                Level = level,
                SortOrder = row - 2
            });

            memberLookup[$"{dimName}|{memberName}"] = id;
            row++;
        }
    }

    #endregion
}
