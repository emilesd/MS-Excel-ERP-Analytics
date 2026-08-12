namespace MyOlap.Core;

/// <summary>
/// Compact serialization for ViewState so layouts can be persisted in SQLite and
/// in a workbook meta sheet (same durability model as other MyOlap settings).
/// Format: v1|{modelId}|R:{dimId}:{id,id}|C:{dimId}:{id,id}|P:{dimId}:{mbrId}|...
/// </summary>
public static class ViewStateCodec
{
    public static long? TryParseModelId(string? payload)
    {
        if (string.IsNullOrWhiteSpace(payload)) return null;
        var parts = payload.Split('|');
        if (parts.Length < 2 || parts[0] != "v1") return null;
        return long.TryParse(parts[1], out var id) ? id : null;
    }

    public static string Serialize(ViewState view)
    {
        var parts = new List<string> { "v1", view.ModelId.ToString() };
        foreach (var a in view.RowAxes)
            parts.Add($"R:{a.DimensionId}:{string.Join(",", a.VisibleMemberIds)}");
        foreach (var a in view.ColAxes)
            parts.Add($"C:{a.DimensionId}:{string.Join(",", a.VisibleMemberIds)}");
        foreach (var kv in view.PovSelections.OrderBy(k => k.Key))
            parts.Add($"P:{kv.Key}:{kv.Value}");
        return string.Join("|", parts);
    }

    public static ViewState? Deserialize(string? payload, Func<long, string?> dimNameResolver)
    {
        if (string.IsNullOrWhiteSpace(payload)) return null;
        var parts = payload.Split('|');
        if (parts.Length < 4 || parts[0] != "v1") return null;
        if (!long.TryParse(parts[1], out var modelId)) return null;

        var view = new ViewState { ModelId = modelId };
        foreach (var part in parts.Skip(2))
        {
            if (part.Length < 3 || part[1] != ':') continue;
            var body = part[2..];
            var colon = body.IndexOf(':');
            if (colon <= 0) continue;
            if (!long.TryParse(body[..colon], out var dimId)) continue;
            var rest = body[(colon + 1)..];

            switch (part[0])
            {
                case 'R':
                case 'C':
                    var ids = rest.Split(',', StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => long.TryParse(s, out var id) ? id : 0)
                        .Where(id => id > 0).ToList();
                    if (ids.Count == 0) continue;
                    var axis = new DimensionAxis
                    {
                        DimensionId = dimId,
                        DimensionName = dimNameResolver(dimId) ?? "",
                        VisibleMemberIds = ids
                    };
                    if (part[0] == 'R') view.RowAxes.Add(axis);
                    else view.ColAxes.Add(axis);
                    break;
                case 'P':
                    if (long.TryParse(rest, out var mbrId) && mbrId > 0)
                        view.PovSelections[dimId] = mbrId;
                    break;
            }
        }

        if (view.RowAxes.Count == 0 || view.ColAxes.Count == 0) return null;
        return view;
    }
}
