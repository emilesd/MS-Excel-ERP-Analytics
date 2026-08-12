using MyOlap.Data;
using MyOlap.Core;

Console.WriteLine("=== Existing Models in DB ===");
try
{
    SqliteRepository.Instance.EnsureDatabaseCreated();
    var models = SqliteRepository.Instance.GetAllModels();
    if (models.Count == 0) {
        Console.WriteLine("No models found.");
    }
    foreach (var m in models) {
        var dims = SqliteRepository.Instance.GetDimensions(m.Id);
        var facts = SqliteRepository.Instance.GetAllFacts(m.Id);
        Console.WriteLine($"  Id={m.Id}  Name='{m.Name}'  Dims={dims.Count}  Facts={facts.Count}");
        foreach (var d in dims) {
            var roots = SqliteRepository.Instance.GetRootMembers(d.Id);
            int totalMembers = 0;
            void Count(long id) {
                var ch = SqliteRepository.Instance.GetChildren(id);
                totalMembers += ch.Count;
                foreach (var c in ch) Count(c.Id);
            }
            foreach (var r in roots) Count(r.Id);
            totalMembers += roots.Count;
            Console.WriteLine($"    Dim: '{d.Name}' ({d.DimType})  roots={roots.Count}  members={totalMembers}");
        }
    }
}
catch (Exception ex) {
    Console.WriteLine($"ERROR: {ex.Message}");
}
