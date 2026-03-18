namespace MyOlap.Data;

/// <summary>
/// Dimension type flags matching the product brief:
/// 5 pre-defined (View, Version, Time, Year, Measure) + up to 7 user-defined.
/// </summary>
public enum DimensionType
{
    UserDefined = 0,
    View = 1,
    Version = 2,
    Time = 3,
    Year = 4,
    Measure = 5
}

/// <summary>
/// Measure/Fact types as specified: Accounting or Statistical.
/// </summary>
public enum MeasureFactType
{
    Accounting = 0,
    Statistical = 1
}

/// <summary>
/// Measure data types: Numeric or Text.
/// </summary>
public enum MeasureDataType
{
    Numeric = 0,
    Text = 1
}

public class OlapModel
{
    public long Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
}

public class Dimension
{
    public long Id { get; set; }
    public long ModelId { get; set; }
    public string Name { get; set; } = string.Empty;
    public DimensionType DimType { get; set; }
    public int SortOrder { get; set; }
}

public class Member
{
    public long Id { get; set; }
    public long DimensionId { get; set; }
    public long? ParentId { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public int Level { get; set; }
    public int SortOrder { get; set; }
}

/// <summary>
/// Filter properties on dimension members (up to 5 per member).
/// E.g. Country, Currency, Location on a Business Unit member.
/// </summary>
public class MemberFilter
{
    public long Id { get; set; }
    public long MemberId { get; set; }
    public string FilterName { get; set; } = string.Empty;
    public string FilterValue { get; set; } = string.Empty;
}

/// <summary>
/// Stores intersection values keyed by a composite member key.
/// The MemberKey is a pipe-delimited string of member IDs, one per dimension.
/// </summary>
public class FactData
{
    public long Id { get; set; }
    public long ModelId { get; set; }
    public string MemberKey { get; set; } = string.Empty;
    public decimal? NumericValue { get; set; }
    public string? TextValue { get; set; }
    public MeasureDataType DataType { get; set; }
}

/// <summary>
/// Per-model user settings persisted across sessions.
/// </summary>
public class ModelSettings
{
    public long ModelId { get; set; }
    public bool OmitEmptyRows { get; set; }
    public bool OmitEmptyColumns { get; set; }
    /// <summary>0 = Name only, 1 = Description only, 2 = Both</summary>
    public int MemberDisplay { get; set; }
}
