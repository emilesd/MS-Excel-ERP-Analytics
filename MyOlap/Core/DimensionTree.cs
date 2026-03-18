using MyOlap.Data;

namespace MyOlap.Core;

/// <summary>
/// In-memory tree node mirroring the hierarchical parent-child structure
/// stored in SQLite. Used for the member-picker tree view and drill operations.
/// </summary>
public class DimensionTreeNode
{
    public Member Member { get; set; } = null!;
    public List<DimensionTreeNode> Children { get; set; } = new();
    public bool IsLeaf => Children.Count == 0;
}

/// <summary>
/// Builds and queries an in-memory tree from the Members table
/// for a given dimension.
/// </summary>
public static class DimensionTree
{
    public static List<DimensionTreeNode> BuildTree(long dimensionId)
    {
        var repo = SqliteRepository.Instance;
        var allMembers = repo.GetMembers(dimensionId);
        var lookup = allMembers.ToDictionary(m => m.Id);
        var roots = new List<DimensionTreeNode>();
        var nodeMap = new Dictionary<long, DimensionTreeNode>();

        foreach (var m in allMembers)
            nodeMap[m.Id] = new DimensionTreeNode { Member = m };

        foreach (var m in allMembers)
        {
            if (m.ParentId is null)
                roots.Add(nodeMap[m.Id]);
            else if (nodeMap.TryGetValue(m.ParentId.Value, out var parent))
                parent.Children.Add(nodeMap[m.Id]);
        }
        return roots;
    }

    /// <summary>
    /// Flattens a tree of nodes into a depth-first list.
    /// </summary>
    public static List<Member> Flatten(IEnumerable<DimensionTreeNode> nodes)
    {
        var result = new List<Member>();
        foreach (var n in nodes)
        {
            result.Add(n.Member);
            result.AddRange(Flatten(n.Children));
        }
        return result;
    }

    /// <summary>
    /// Finds a node by member ID in a tree.
    /// </summary>
    public static DimensionTreeNode? Find(IEnumerable<DimensionTreeNode> roots, long memberId)
    {
        foreach (var root in roots)
        {
            if (root.Member.Id == memberId) return root;
            var found = Find(root.Children, memberId);
            if (found != null) return found;
        }
        return null;
    }
}
