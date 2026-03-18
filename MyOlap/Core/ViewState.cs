using MyOlap.Data;

namespace MyOlap.Core;

/// <summary>
/// Captures the full state of a worksheet view at a point in time,
/// used by the undo stack and the OLAP engine to render/restore grids.
/// </summary>
public class ViewState
{
    public long ModelId { get; set; }

    /// <summary>Dimensions placed on the row axis (left-to-right heading columns).</summary>
    public List<DimensionAxis> RowAxes { get; set; } = new();

    /// <summary>Dimensions placed on the column axis (top heading rows).</summary>
    public List<DimensionAxis> ColAxes { get; set; } = new();

    /// <summary>Point-of-view: single member selected for each dimension not on row/col.</summary>
    public Dictionary<long, long> PovSelections { get; set; } = new();

    public ViewState Clone()
    {
        return new ViewState
        {
            ModelId = ModelId,
            RowAxes = RowAxes.Select(a => a.Clone()).ToList(),
            ColAxes = ColAxes.Select(a => a.Clone()).ToList(),
            PovSelections = new Dictionary<long, long>(PovSelections)
        };
    }
}

/// <summary>
/// One dimension on an axis with the list of members currently expanded/visible.
/// </summary>
public class DimensionAxis
{
    public long DimensionId { get; set; }
    public string DimensionName { get; set; } = string.Empty;
    public List<long> VisibleMemberIds { get; set; } = new();

    public DimensionAxis Clone()
    {
        return new DimensionAxis
        {
            DimensionId = DimensionId,
            DimensionName = DimensionName,
            VisibleMemberIds = new List<long>(VisibleMemberIds)
        };
    }
}
