using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Data.Sqlite;

namespace MyOlap.Data;

/// <summary>
/// Singleton data-access layer backed by a local SQLite file stored next to the add-in.
/// All models, dimensions, members, filters, fact data, and settings live here.
/// </summary>
public sealed class SqliteRepository : IDisposable
{
    private static readonly Lazy<SqliteRepository> _lazy = new(() => new SqliteRepository());
    public static SqliteRepository Instance => _lazy.Value;
    private static bool _batteriesInitialized;

    private readonly string _dbPath;
    private SqliteConnection? _conn;

    private SqliteRepository()
    {
        InitBatteries();
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MyOlap");
        Directory.CreateDirectory(dir);
        _dbPath = Path.Combine(dir, "myolap.db");
    }

    private static void InitBatteries()
    {
        if (_batteriesInitialized) return;
        try
        {
            RegisterNativeResolver();
            SQLitePCL.Batteries_V2.Init();
            _batteriesInitialized = true;
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"SQLite init error:\n{ex.Message}\n\nInner: {ex.InnerException?.Message}\n\nStack: {ex.StackTrace}",
                "MyOlap SQLite Init", System.Windows.Forms.MessageBoxButtons.OK);
        }
    }

    /// <summary>
    /// Registers a native library resolver that searches the add-in's directory
    /// and the runtimes subdirectory for e_sqlite3.dll.
    /// </summary>
    private static void RegisterNativeResolver()
    {
        NativeLibrary.SetDllImportResolver(
            typeof(SQLitePCL.raw).Assembly,
            (libraryName, assembly, searchPath) =>
            {
                if (!libraryName.Contains("e_sqlite3"))
                    return IntPtr.Zero;

                var candidates = new List<string>();

                var asmDir = Path.GetDirectoryName(assembly.Location) ?? "";
                candidates.Add(Path.Combine(asmDir, "e_sqlite3.dll"));
                candidates.Add(Path.Combine(asmDir, "runtimes", "win-x64", "native", "e_sqlite3.dll"));

                var exeDir = Path.GetDirectoryName(typeof(SqliteRepository).Assembly.Location) ?? "";
                candidates.Add(Path.Combine(exeDir, "e_sqlite3.dll"));
                candidates.Add(Path.Combine(exeDir, "runtimes", "win-x64", "native", "e_sqlite3.dll"));

                var baseDir = AppContext.BaseDirectory;
                candidates.Add(Path.Combine(baseDir, "e_sqlite3.dll"));
                candidates.Add(Path.Combine(baseDir, "runtimes", "win-x64", "native", "e_sqlite3.dll"));

                foreach (var path in candidates)
                {
                    if (File.Exists(path) && NativeLibrary.TryLoad(path, out var handle))
                        return handle;
                }

                return IntPtr.Zero;
            });
    }

    private SqliteConnection GetConnection()
    {
        if (_conn is { State: System.Data.ConnectionState.Open })
            return _conn;
        _conn = new SqliteConnection($"Data Source={_dbPath}");
        _conn.Open();
        using var cmd = _conn.CreateCommand();
        cmd.CommandText = "PRAGMA journal_mode=WAL; PRAGMA foreign_keys=ON;";
        cmd.ExecuteNonQuery();
        return _conn;
    }

    public void EnsureDatabaseCreated()
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS Models (
    Id          INTEGER PRIMARY KEY AUTOINCREMENT,
    Name        TEXT    NOT NULL,
    Description TEXT    NOT NULL DEFAULT '',
    CreatedUtc  TEXT    NOT NULL
);

CREATE TABLE IF NOT EXISTS Dimensions (
    Id        INTEGER PRIMARY KEY AUTOINCREMENT,
    ModelId   INTEGER NOT NULL REFERENCES Models(Id) ON DELETE CASCADE,
    Name      TEXT    NOT NULL,
    DimType   INTEGER NOT NULL DEFAULT 0,
    SortOrder INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS Members (
    Id          INTEGER PRIMARY KEY AUTOINCREMENT,
    DimensionId INTEGER NOT NULL REFERENCES Dimensions(Id) ON DELETE CASCADE,
    ParentId    INTEGER REFERENCES Members(Id) ON DELETE SET NULL,
    Name        TEXT    NOT NULL,
    Description TEXT    NOT NULL DEFAULT '',
    Level       INTEGER NOT NULL DEFAULT 0,
    SortOrder   INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS MemberFilters (
    Id          INTEGER PRIMARY KEY AUTOINCREMENT,
    MemberId    INTEGER NOT NULL REFERENCES Members(Id) ON DELETE CASCADE,
    FilterName  TEXT    NOT NULL,
    FilterValue TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS FactData (
    Id           INTEGER PRIMARY KEY AUTOINCREMENT,
    ModelId      INTEGER NOT NULL REFERENCES Models(Id) ON DELETE CASCADE,
    MemberKey    TEXT    NOT NULL,
    NumericValue REAL,
    TextValue    TEXT,
    DataType     INTEGER NOT NULL DEFAULT 0
);
CREATE INDEX IF NOT EXISTS IX_FactData_Key ON FactData(ModelId, MemberKey);

CREATE TABLE IF NOT EXISTS ModelSettings (
    ModelId          INTEGER PRIMARY KEY REFERENCES Models(Id) ON DELETE CASCADE,
    OmitEmptyRows    INTEGER NOT NULL DEFAULT 0,
    OmitEmptyColumns INTEGER NOT NULL DEFAULT 0,
    MemberDisplay    INTEGER NOT NULL DEFAULT 0
);
";
        cmd.ExecuteNonQuery();
    }

    #region Models

    public long InsertModel(OlapModel model)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "INSERT INTO Models (Name, Description, CreatedUtc) VALUES ($n, $d, $c); SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("$n", model.Name);
        cmd.Parameters.AddWithValue("$d", model.Description);
        cmd.Parameters.AddWithValue("$c", model.CreatedUtc.ToString("o"));
        return (long)cmd.ExecuteScalar()!;
    }

    public List<OlapModel> GetAllModels()
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Name, Description, CreatedUtc FROM Models ORDER BY Name";
        using var rdr = cmd.ExecuteReader();
        var list = new List<OlapModel>();
        while (rdr.Read())
        {
            list.Add(new OlapModel
            {
                Id = rdr.GetInt64(0),
                Name = rdr.GetString(1),
                Description = rdr.GetString(2),
                CreatedUtc = DateTime.Parse(rdr.GetString(3))
            });
        }
        return list;
    }

    public void DeleteModel(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Models WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", modelId);
        cmd.ExecuteNonQuery();
    }

    #endregion

    #region Dimensions

    public long InsertDimension(Dimension dim)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "INSERT INTO Dimensions (ModelId, Name, DimType, SortOrder) VALUES ($m, $n, $t, $s); SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("$m", dim.ModelId);
        cmd.Parameters.AddWithValue("$n", dim.Name);
        cmd.Parameters.AddWithValue("$t", (int)dim.DimType);
        cmd.Parameters.AddWithValue("$s", dim.SortOrder);
        return (long)cmd.ExecuteScalar()!;
    }

    public List<Dimension> GetDimensions(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, ModelId, Name, DimType, SortOrder FROM Dimensions WHERE ModelId = $m ORDER BY SortOrder";
        cmd.Parameters.AddWithValue("$m", modelId);
        using var rdr = cmd.ExecuteReader();
        var list = new List<Dimension>();
        while (rdr.Read())
        {
            list.Add(new Dimension
            {
                Id = rdr.GetInt64(0),
                ModelId = rdr.GetInt64(1),
                Name = rdr.GetString(2),
                DimType = (DimensionType)rdr.GetInt32(3),
                SortOrder = rdr.GetInt32(4)
            });
        }
        return list;
    }

    public void UpdateDimension(Dimension dim)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "UPDATE Dimensions SET Name = $n, DimType = $t, SortOrder = $s WHERE Id = $id";
        cmd.Parameters.AddWithValue("$n", dim.Name);
        cmd.Parameters.AddWithValue("$t", (int)dim.DimType);
        cmd.Parameters.AddWithValue("$s", dim.SortOrder);
        cmd.Parameters.AddWithValue("$id", dim.Id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteDimension(long dimId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Dimensions WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", dimId);
        cmd.ExecuteNonQuery();
    }

    #endregion

    #region Members

    public long InsertMember(Member member)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO Members (DimensionId, ParentId, Name, Description, Level, SortOrder)
VALUES ($d, $p, $n, $desc, $l, $s);
SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("$d", member.DimensionId);
        cmd.Parameters.AddWithValue("$p", (object?)member.ParentId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("$n", member.Name);
        cmd.Parameters.AddWithValue("$desc", member.Description);
        cmd.Parameters.AddWithValue("$l", member.Level);
        cmd.Parameters.AddWithValue("$s", member.SortOrder);
        return (long)cmd.ExecuteScalar()!;
    }

    public List<Member> GetMembers(long dimensionId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder FROM Members WHERE DimensionId = $d ORDER BY SortOrder";
        cmd.Parameters.AddWithValue("$d", dimensionId);
        using var rdr = cmd.ExecuteReader();
        var list = new List<Member>();
        while (rdr.Read())
        {
            list.Add(new Member
            {
                Id = rdr.GetInt64(0),
                DimensionId = rdr.GetInt64(1),
                ParentId = rdr.IsDBNull(2) ? null : rdr.GetInt64(2),
                Name = rdr.GetString(3),
                Description = rdr.GetString(4),
                Level = rdr.GetInt32(5),
                SortOrder = rdr.GetInt32(6)
            });
        }
        return list;
    }

    public List<Member> GetChildren(long parentId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder FROM Members WHERE ParentId = $p ORDER BY SortOrder";
        cmd.Parameters.AddWithValue("$p", parentId);
        using var rdr = cmd.ExecuteReader();
        var list = new List<Member>();
        while (rdr.Read())
        {
            list.Add(new Member
            {
                Id = rdr.GetInt64(0),
                DimensionId = rdr.GetInt64(1),
                ParentId = rdr.IsDBNull(2) ? null : rdr.GetInt64(2),
                Name = rdr.GetString(3),
                Description = rdr.GetString(4),
                Level = rdr.GetInt32(5),
                SortOrder = rdr.GetInt32(6)
            });
        }
        return list;
    }

    public List<Member> GetRootMembers(long dimensionId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder FROM Members WHERE DimensionId = $d AND ParentId IS NULL ORDER BY SortOrder";
        cmd.Parameters.AddWithValue("$d", dimensionId);
        using var rdr = cmd.ExecuteReader();
        var list = new List<Member>();
        while (rdr.Read())
        {
            list.Add(new Member
            {
                Id = rdr.GetInt64(0),
                DimensionId = rdr.GetInt64(1),
                ParentId = null,
                Name = rdr.GetString(3),
                Description = rdr.GetString(4),
                Level = rdr.GetInt32(5),
                SortOrder = rdr.GetInt32(6)
            });
        }
        return list;
    }

    public List<Member> GetAllDescendants(long memberId)
    {
        var result = new List<Member>();
        var children = GetChildren(memberId);
        foreach (var child in children)
        {
            result.Add(child);
            result.AddRange(GetAllDescendants(child.Id));
        }
        return result;
    }

    public List<Member> GetLeafDescendants(long memberId)
    {
        var all = GetAllDescendants(memberId);
        return all.Where(m => GetChildren(m.Id).Count == 0).ToList();
    }

    public Member? GetMember(long memberId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder FROM Members WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", memberId);
        using var rdr = cmd.ExecuteReader();
        if (!rdr.Read()) return null;
        return new Member
        {
            Id = rdr.GetInt64(0),
            DimensionId = rdr.GetInt64(1),
            ParentId = rdr.IsDBNull(2) ? null : rdr.GetInt64(2),
            Name = rdr.GetString(3),
            Description = rdr.GetString(4),
            Level = rdr.GetInt32(5),
            SortOrder = rdr.GetInt32(6)
        };
    }

    public void UpdateMember(Member m)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "UPDATE Members SET ParentId=$p, Name=$n, Description=$desc, Level=$l, SortOrder=$s WHERE Id=$id";
        cmd.Parameters.AddWithValue("$p", (object?)m.ParentId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("$n", m.Name);
        cmd.Parameters.AddWithValue("$desc", m.Description);
        cmd.Parameters.AddWithValue("$l", m.Level);
        cmd.Parameters.AddWithValue("$s", m.SortOrder);
        cmd.Parameters.AddWithValue("$id", m.Id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteMember(long memberId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Members WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", memberId);
        cmd.ExecuteNonQuery();
    }

    #endregion

    #region MemberFilters

    public void InsertFilter(MemberFilter f)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "INSERT INTO MemberFilters (MemberId, FilterName, FilterValue) VALUES ($m, $n, $v)";
        cmd.Parameters.AddWithValue("$m", f.MemberId);
        cmd.Parameters.AddWithValue("$n", f.FilterName);
        cmd.Parameters.AddWithValue("$v", f.FilterValue);
        cmd.ExecuteNonQuery();
    }

    public List<MemberFilter> GetFilters(long memberId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, MemberId, FilterName, FilterValue FROM MemberFilters WHERE MemberId = $m";
        cmd.Parameters.AddWithValue("$m", memberId);
        using var rdr = cmd.ExecuteReader();
        var list = new List<MemberFilter>();
        while (rdr.Read())
        {
            list.Add(new MemberFilter
            {
                Id = rdr.GetInt64(0),
                MemberId = rdr.GetInt64(1),
                FilterName = rdr.GetString(2),
                FilterValue = rdr.GetString(3)
            });
        }
        return list;
    }

    public void DeleteFilters(long memberId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM MemberFilters WHERE MemberId = $m";
        cmd.Parameters.AddWithValue("$m", memberId);
        cmd.ExecuteNonQuery();
    }

    #endregion

    #region FactData

    public void InsertFactBatch(long modelId, IEnumerable<FactData> facts)
    {
        var conn = GetConnection();
        using var tx = conn.BeginTransaction();
        using var cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = @"
INSERT OR REPLACE INTO FactData (ModelId, MemberKey, NumericValue, TextValue, DataType)
VALUES ($m, $k, $nv, $tv, $dt)";
        var pM = cmd.Parameters.Add("$m", SqliteType.Integer);
        var pK = cmd.Parameters.Add("$k", SqliteType.Text);
        var pNv = cmd.Parameters.Add("$nv", SqliteType.Real);
        var pTv = cmd.Parameters.Add("$tv", SqliteType.Text);
        var pDt = cmd.Parameters.Add("$dt", SqliteType.Integer);
        foreach (var f in facts)
        {
            pM.Value = modelId;
            pK.Value = f.MemberKey;
            pNv.Value = (object?)f.NumericValue ?? DBNull.Value;
            pTv.Value = (object?)f.TextValue ?? DBNull.Value;
            pDt.Value = (int)f.DataType;
            cmd.ExecuteNonQuery();
        }
        tx.Commit();
    }

    public decimal? GetFactValue(long modelId, string memberKey)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT NumericValue FROM FactData WHERE ModelId = $m AND MemberKey = $k";
        cmd.Parameters.AddWithValue("$m", modelId);
        cmd.Parameters.AddWithValue("$k", memberKey);
        var result = cmd.ExecuteScalar();
        if (result is null or DBNull) return null;
        return Convert.ToDecimal(result);
    }

    public FactData? GetFact(long modelId, string memberKey)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, ModelId, MemberKey, NumericValue, TextValue, DataType FROM FactData WHERE ModelId = $m AND MemberKey = $k";
        cmd.Parameters.AddWithValue("$m", modelId);
        cmd.Parameters.AddWithValue("$k", memberKey);
        using var rdr = cmd.ExecuteReader();
        if (!rdr.Read()) return null;
        return new FactData
        {
            Id = rdr.GetInt64(0),
            ModelId = rdr.GetInt64(1),
            MemberKey = rdr.GetString(2),
            NumericValue = rdr.IsDBNull(3) ? null : (decimal)rdr.GetDouble(3),
            TextValue = rdr.IsDBNull(4) ? null : rdr.GetString(4),
            DataType = (MeasureDataType)rdr.GetInt32(5)
        };
    }

    public void ClearFacts(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM FactData WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        cmd.ExecuteNonQuery();
    }

    #endregion

    #region Settings

    public ModelSettings GetSettings(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT ModelId, OmitEmptyRows, OmitEmptyColumns, MemberDisplay FROM ModelSettings WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        using var rdr = cmd.ExecuteReader();
        if (rdr.Read())
        {
            return new ModelSettings
            {
                ModelId = rdr.GetInt64(0),
                OmitEmptyRows = rdr.GetInt32(1) != 0,
                OmitEmptyColumns = rdr.GetInt32(2) != 0,
                MemberDisplay = rdr.GetInt32(3)
            };
        }
        return new ModelSettings { ModelId = modelId };
    }

    public void SaveSettings(ModelSettings s)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO ModelSettings (ModelId, OmitEmptyRows, OmitEmptyColumns, MemberDisplay)
VALUES ($m, $r, $c, $d)
ON CONFLICT(ModelId) DO UPDATE SET
    OmitEmptyRows = excluded.OmitEmptyRows,
    OmitEmptyColumns = excluded.OmitEmptyColumns,
    MemberDisplay = excluded.MemberDisplay";
        cmd.Parameters.AddWithValue("$m", s.ModelId);
        cmd.Parameters.AddWithValue("$r", s.OmitEmptyRows ? 1 : 0);
        cmd.Parameters.AddWithValue("$c", s.OmitEmptyColumns ? 1 : 0);
        cmd.Parameters.AddWithValue("$d", s.MemberDisplay);
        cmd.ExecuteNonQuery();
    }

    #endregion

    public void Dispose()
    {
        _conn?.Dispose();
        _conn = null;
    }
}
