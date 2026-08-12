namespace MyOlap.Core;

/// <summary>
/// Parses MyOlap metadata embedded in Excel cell comments/notes
/// (MODEL:x, DIM:x, MBR:y). Tolerates Excel author-name prefixes.
/// </summary>
public static class SheetMetaParser
{
    public static string? ExtractDimMbrKey(string? commentText)
    {
        if (string.IsNullOrEmpty(commentText)) return null;
        string? dimSeg = null, mbrSeg = null;
        foreach (var line in commentText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            foreach (var seg in line.Split('|'))
            {
                var s = seg.Trim();
                if (s.StartsWith("DIM:")) dimSeg = s;
                else if (s.StartsWith("MBR:")) mbrSeg = s;
            }
            if (dimSeg != null && mbrSeg != null) break;
        }
        return (dimSeg != null && mbrSeg != null) ? $"{dimSeg}|{mbrSeg}" : null;
    }

    public static long? ExtractModelId(string? commentText)
    {
        if (string.IsNullOrEmpty(commentText)) return null;
        foreach (var line in commentText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            foreach (var seg in line.Split('|'))
            {
                var s = seg.Trim();
                if (s.StartsWith("MODEL:") && long.TryParse(s.AsSpan(6), out var id))
                    return id;
            }
        return null;
    }

    public static bool TryParseDimMbr(string? key, out long dimId, out long mbrId)
    {
        dimId = 0;
        mbrId = 0;
        if (string.IsNullOrEmpty(key)) return false;
        foreach (var seg in key.Split('|'))
        {
            var s = seg.Trim();
            // "DIM:" and "MBR:" are both 4 chars — must use [4..] (not [5..]).
            if (s.StartsWith("DIM:")) long.TryParse(s.AsSpan(4), out dimId);
            else if (s.StartsWith("MBR:")) long.TryParse(s.AsSpan(4), out mbrId);
        }
        return dimId != 0 && mbrId != 0;
    }
}
