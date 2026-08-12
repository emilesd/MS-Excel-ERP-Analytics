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

    // In-memory caches — invalidated on relevant mutations.
    private readonly Dictionary<long, List<Member>> _childrenCache = new();
    private readonly Dictionary<long, List<Member>> _membersCache = new();
    private readonly Dictionary<long, Member?> _memberCache = new();
    private readonly Dictionary<long, List<Dimension>> _dimensionsCache = new();
    private readonly Dictionary<long, ModelSettings> _settingsCache = new();

    public void InvalidateMemberCache()
    {
        _childrenCache.Clear();
        _membersCache.Clear();
        _memberCache.Clear();
    }

    private void InvalidateDimensionCache(long modelId)
    {
        _dimensionsCache.Remove(modelId);
    }

    private void InvalidateSettingsCache(long modelId)
    {
        _settingsCache.Remove(modelId);
    }

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
            // ExcelDNA packed XLL (LoadFromBytes=true) does NOT extract native DLLs to disk.
            // Register a resolver on the provider assembly so P/Invoke finds e_sqlite3.dll
            // from the AddIn directory where the XLL lives.
            NativeLibrary.SetDllImportResolver(
                typeof(SQLitePCL.SQLite3Provider_e_sqlite3).Assembly,
                ResolveNativeSqlite3);

            SQLitePCL.Batteries_V2.Init();
            _batteriesInitialized = true;
        }
        catch (InvalidOperationException)
        {
            // Batteries_V2.Init() throws if the [ModuleInitializer] already set the provider.
            _batteriesInitialized = true;
        }
        catch (Exception ex)
        {
            System.Windows.Forms.MessageBox.Show(
                $"SQLite init error:\n{ex.Message}\n\nInner: {ex.InnerException?.Message}",
                "MyOlap SQLite Init", System.Windows.Forms.MessageBoxButtons.OK);
        }
    }

    private static IntPtr ResolveNativeSqlite3(string libraryName, Assembly asm, DllImportSearchPath? searchPath)
    {
        if (!libraryName.Contains("e_sqlite3")) return IntPtr.Zero;

        var candidates = new List<string>();

        // AddIn directory (where the XLL is registered) — primary location
        try
        {
            var xllPath = ExcelDna.Integration.ExcelDnaUtil.XllPath;
            var xllDir = Path.GetDirectoryName(xllPath) ?? "";
            if (!string.IsNullOrEmpty(xllDir))
            {
                candidates.Add(Path.Combine(xllDir, "e_sqlite3.dll"));
                candidates.Add(Path.Combine(xllDir, "runtimes", "win-x64", "native", "e_sqlite3.dll"));
            }
        }
        catch { }

        // Fallback: assembly location (valid when NOT packed / in TestConsole)
        var asmDir = Path.GetDirectoryName(asm.Location) ?? "";
        if (!string.IsNullOrEmpty(asmDir))
        {
            candidates.Add(Path.Combine(asmDir, "e_sqlite3.dll"));
            candidates.Add(Path.Combine(asmDir, "runtimes", "win-x64", "native", "e_sqlite3.dll"));
        }

        foreach (var path in candidates)
        {
            if (File.Exists(path) && NativeLibrary.TryLoad(path, out var handle))
                return handle;
        }
        return IntPtr.Zero;
    }

    private SqliteConnection GetConnection()
    {
        if (_conn is { State: System.Data.ConnectionState.Open })
            return _conn;

        var csb = new Microsoft.Data.Sqlite.SqliteConnectionStringBuilder { DataSource = _dbPath };
        _conn = new SqliteConnection(csb.ToString());
        try
        {
            _conn.Open();
        }
        catch
        {
            // Open failed — existing DB is likely the old SQLCipher-encrypted file.
            // Rename it and start fresh; user will need to re-import data.
            _conn.Dispose();
            _conn = null;
            if (File.Exists(_dbPath))
                File.Move(_dbPath, _dbPath + ".encrypted_backup", overwrite: true);
            _conn = new SqliteConnection(csb.ToString());
            _conn.Open();
        }
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
    SortOrder       INTEGER NOT NULL DEFAULT 0,
    ConsolOperator TEXT    NOT NULL DEFAULT '+' 
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
        using var migCmd = conn.CreateCommand();
        migCmd.CommandText = "SELECT COUNT(*) FROM pragma_table_info('Members') WHERE name='ConsolOperator'";
        var hasCol = Convert.ToInt64(migCmd.ExecuteScalar()) > 0;
        if (!hasCol)
        {
            using var alterCmd = conn.CreateCommand();
            alterCmd.CommandText = "ALTER TABLE Members ADD COLUMN ConsolOperator TEXT NOT NULL DEFAULT '+'";
            alterCmd.ExecuteNonQuery();
        }

        MigrateColumn(conn, "Members", "Formula", "TEXT NOT NULL DEFAULT ''");
        MigrateColumn(conn, "Members", "TimeBalance", "TEXT NOT NULL DEFAULT ''");
        MigrateColumn(conn, "Members", "SharedFromId", "INTEGER REFERENCES Members(Id)");
        MigrateColumn(conn, "ModelSettings", "PreserveFormulas", "INTEGER NOT NULL DEFAULT 0");

        // Last-used OLAP layout per model (survives Excel restart; same durability as settings).
        using (var viewCmd = conn.CreateCommand())
        {
            viewCmd.CommandText = @"
CREATE TABLE IF NOT EXISTS ModelViews (
    ModelId     INTEGER PRIMARY KEY REFERENCES Models(Id) ON DELETE CASCADE,
    ViewPayload TEXT NOT NULL DEFAULT '',
    UpdatedUtc  TEXT NOT NULL DEFAULT ''
);";
            viewCmd.ExecuteNonQuery();
        }
    }

    private static void MigrateColumn(SqliteConnection conn, string table, string column, string definition)
    {
        using var chk = conn.CreateCommand();
        chk.CommandText = $"SELECT COUNT(*) FROM pragma_table_info('{table}') WHERE name='{column}'";
        if (Convert.ToInt64(chk.ExecuteScalar()) > 0) return;
        using var alt = conn.CreateCommand();
        alt.CommandText = $"ALTER TABLE {table} ADD COLUMN {column} {definition}";
        alt.ExecuteNonQuery();
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
        var id = (long)cmd.ExecuteScalar()!;
        InvalidateDimensionCache(dim.ModelId);
        return id;
    }

    public List<Dimension> GetDimensions(long modelId)
    {
        if (_dimensionsCache.TryGetValue(modelId, out var cached)) return cached;
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
        _dimensionsCache[modelId] = list;
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
        InvalidateDimensionCache(dim.ModelId);
    }

    public void DeleteDimension(long dimId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        // Fetch modelId before deleting so we can invalidate the right cache entry.
        using var qry = conn.CreateCommand();
        qry.CommandText = "SELECT ModelId FROM Dimensions WHERE Id = $id";
        qry.Parameters.AddWithValue("$id", dimId);
        var modelId = qry.ExecuteScalar() is long mid ? mid : -1L;
        cmd.CommandText = "DELETE FROM Dimensions WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", dimId);
        cmd.ExecuteNonQuery();
        if (modelId >= 0) InvalidateDimensionCache(modelId);
    }

    #endregion

    #region Members

    public long InsertMember(Member member)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO Members (DimensionId, ParentId, Name, Description, Level, SortOrder, ConsolOperator, Formula, TimeBalance, SharedFromId)
VALUES ($d, $p, $n, $desc, $l, $s, $co, $f, $tb, $sf);
SELECT last_insert_rowid();";
        cmd.Parameters.AddWithValue("$d", member.DimensionId);
        cmd.Parameters.AddWithValue("$p", (object?)member.ParentId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("$n", member.Name);
        cmd.Parameters.AddWithValue("$desc", member.Description);
        cmd.Parameters.AddWithValue("$l", member.Level);
        cmd.Parameters.AddWithValue("$s", member.SortOrder);
        cmd.Parameters.AddWithValue("$co", member.ConsolOperator ?? "+");
        cmd.Parameters.AddWithValue("$f", member.Formula ?? "");
        cmd.Parameters.AddWithValue("$tb", member.TimeBalance ?? "");
        cmd.Parameters.AddWithValue("$sf", (object?)member.SharedFromId ?? DBNull.Value);
        var id = (long)cmd.ExecuteScalar()!;
        InvalidateMemberCache();
        return id;
    }

    public List<Member> GetMembers(long dimensionId)
    {
        if (_membersCache.TryGetValue(dimensionId, out var hit)) return hit;
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder, ConsolOperator, Formula, TimeBalance, SharedFromId FROM Members WHERE DimensionId = $d ORDER BY SortOrder";
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
                SortOrder = rdr.GetInt32(6),
                ConsolOperator = rdr.IsDBNull(7) ? "+" : rdr.GetString(7),
                Formula = rdr.IsDBNull(8) ? "" : rdr.GetString(8),
                TimeBalance = rdr.IsDBNull(9) ? "" : rdr.GetString(9),
                SharedFromId = rdr.IsDBNull(10) ? null : rdr.GetInt64(10)
            });
        }
        _membersCache[dimensionId] = list;

        // Pre-warm per-member caches so GetMember() and GetChildren() never hit the DB
        // for any member in this dimension after this single bulk load.
        foreach (var m in list)
            _memberCache[m.Id] = m;

        var childMap = new Dictionary<long, List<Member>>();
        foreach (var m in list)
        {
            var key = m.ParentId ?? 0L;
            if (!childMap.TryGetValue(key, out var siblings))
                childMap[key] = siblings = new List<Member>();
            siblings.Add(m);
        }
        foreach (var m in list)
            _childrenCache[m.Id] = childMap.TryGetValue(m.Id, out var ch) ? ch : new List<Member>();

        return list;
    }

    public List<Member> GetChildren(long parentId)
    {
        if (_childrenCache.TryGetValue(parentId, out var hit)) return hit;
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder, ConsolOperator, Formula, TimeBalance, SharedFromId FROM Members WHERE ParentId = $p ORDER BY SortOrder";
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
                SortOrder = rdr.GetInt32(6),
                ConsolOperator = rdr.IsDBNull(7) ? "+" : rdr.GetString(7),
                Formula = rdr.IsDBNull(8) ? "" : rdr.GetString(8),
                TimeBalance = rdr.IsDBNull(9) ? "" : rdr.GetString(9),
                SharedFromId = rdr.IsDBNull(10) ? null : rdr.GetInt64(10)
            });
        }
        _childrenCache[parentId] = list;
        return list;
    }

    public List<Member> GetRootMembers(long dimensionId)
    {
        // GetMembers is already cached — filter in memory instead of a separate DB query.
        return GetMembers(dimensionId).Where(m => m.ParentId == null).ToList();
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
        if (_memberCache.TryGetValue(memberId, out var hit)) return hit;
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, DimensionId, ParentId, Name, Description, Level, SortOrder, ConsolOperator, Formula, TimeBalance, SharedFromId FROM Members WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", memberId);
        using var rdr = cmd.ExecuteReader();
        if (!rdr.Read()) { _memberCache[memberId] = null; return null; }
        var m = new Member
        {
            Id = rdr.GetInt64(0),
            DimensionId = rdr.GetInt64(1),
            ParentId = rdr.IsDBNull(2) ? null : rdr.GetInt64(2),
            Name = rdr.GetString(3),
            Description = rdr.GetString(4),
            Level = rdr.GetInt32(5),
            SortOrder = rdr.GetInt32(6),
            ConsolOperator = rdr.IsDBNull(7) ? "+" : rdr.GetString(7),
            Formula = rdr.IsDBNull(8) ? "" : rdr.GetString(8),
            TimeBalance = rdr.IsDBNull(9) ? "" : rdr.GetString(9),
            SharedFromId = rdr.IsDBNull(10) ? null : rdr.GetInt64(10)
        };
        _memberCache[memberId] = m;
        return m;
    }

    public void UpdateMember(Member m)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "UPDATE Members SET ParentId=$p, Name=$n, Description=$desc, Level=$l, SortOrder=$s, ConsolOperator=$co, Formula=$f, TimeBalance=$tb, SharedFromId=$sf WHERE Id=$id";
        cmd.Parameters.AddWithValue("$p", (object?)m.ParentId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("$n", m.Name);
        cmd.Parameters.AddWithValue("$desc", m.Description);
        cmd.Parameters.AddWithValue("$l", m.Level);
        cmd.Parameters.AddWithValue("$s", m.SortOrder);
        cmd.Parameters.AddWithValue("$co", m.ConsolOperator ?? "+");
        cmd.Parameters.AddWithValue("$f", m.Formula ?? "");
        cmd.Parameters.AddWithValue("$tb", m.TimeBalance ?? "");
        cmd.Parameters.AddWithValue("$sf", (object?)m.SharedFromId ?? DBNull.Value);
        cmd.Parameters.AddWithValue("$id", m.Id);
        cmd.ExecuteNonQuery();
        InvalidateMemberCache();
    }

    public Dictionary<string, Member> GetMembersByNameForDimension(long dimensionId)
    {
        var result = new Dictionary<string, Member>(StringComparer.OrdinalIgnoreCase);
        var members = GetMembers(dimensionId);
        foreach (var m in members)
        {
            if (!result.ContainsKey(m.Name))
                result[m.Name] = m;
        }
        return result;
    }

    public void ClearDimensionMembers(long dimensionId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Members WHERE DimensionId = $d";
        cmd.Parameters.AddWithValue("$d", dimensionId);
        cmd.ExecuteNonQuery();
        InvalidateMemberCache();
    }

    public void DeleteMember(long memberId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM Members WHERE Id = $id";
        cmd.Parameters.AddWithValue("$id", memberId);
        cmd.ExecuteNonQuery();
        InvalidateMemberCache();
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

    public List<FactData> GetAllFacts(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT Id, ModelId, MemberKey, NumericValue, TextValue, DataType FROM FactData WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        using var rdr = cmd.ExecuteReader();
        var list = new List<FactData>();
        while (rdr.Read())
        {
            list.Add(new FactData
            {
                Id = rdr.GetInt64(0),
                ModelId = rdr.GetInt64(1),
                MemberKey = rdr.GetString(2),
                NumericValue = rdr.IsDBNull(3) ? null : (decimal)rdr.GetDouble(3),
                TextValue = rdr.IsDBNull(4) ? null : rdr.GetString(4),
                DataType = (MeasureDataType)rdr.GetInt32(5)
            });
        }
        return list;
    }

    public void ClearFactsByFilter(long modelId, List<long> dimOrder, long? viewMemberId, long? versionMemberId, long? yearMemberId, long? viewDimId, long? versionDimId, long? yearDimId)
    {
        var conn = GetConnection();
        var allFacts = GetAllFacts(modelId);

        var viewIds = BuildMemberSet(viewMemberId);
        var versionIds = BuildMemberSet(versionMemberId);
        var yearIds = BuildMemberSet(yearMemberId);

        using var tx = conn.BeginTransaction();
        foreach (var f in allFacts)
        {
            var parts = f.MemberKey.Split('|');
            var memberMap = new Dictionary<long, long>();
            for (int i = 0; i < dimOrder.Count && i < parts.Length; i++)
            {
                if (long.TryParse(parts[i], out var mid) && mid > 0)
                    memberMap[dimOrder[i]] = mid;
            }

            bool matchView = viewDimId == null || viewIds == null || (memberMap.TryGetValue(viewDimId.Value, out var vm) && viewIds.Contains(vm));
            bool matchVersion = versionDimId == null || versionIds == null || (memberMap.TryGetValue(versionDimId.Value, out var vvm) && versionIds.Contains(vvm));
            bool matchYear = yearDimId == null || yearIds == null || (memberMap.TryGetValue(yearDimId.Value, out var ym) && yearIds.Contains(ym));

            if (matchView && matchVersion && matchYear)
            {
                using var cmd = conn.CreateCommand();
                cmd.Transaction = tx;
                cmd.CommandText = "DELETE FROM FactData WHERE Id = @id";
                cmd.Parameters.AddWithValue("@id", f.Id);
                cmd.ExecuteNonQuery();
            }
        }
        tx.Commit();
    }

    private HashSet<long>? BuildMemberSet(long? memberId)
    {
        if (!memberId.HasValue) return null;
        var set = new HashSet<long> { memberId.Value };
        foreach (var desc in GetAllDescendants(memberId.Value))
            set.Add(desc.Id);
        return set;
    }
    public void ClearFacts(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM FactData WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        cmd.ExecuteNonQuery();
    }

    /// <summary>True when the model has at least one fact row (not structure-empty).</summary>
    public bool HasFactData(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT 1 FROM FactData WHERE ModelId = $m LIMIT 1";
        cmd.Parameters.AddWithValue("$m", modelId);
        return cmd.ExecuteScalar() != null;
    }

    #endregion

    #region Settings

    public ModelSettings GetSettings(long modelId)
    {
        if (_settingsCache.TryGetValue(modelId, out var cached)) return cached;
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT ModelId, OmitEmptyRows, OmitEmptyColumns, MemberDisplay, PreserveFormulas FROM ModelSettings WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        using var rdr = cmd.ExecuteReader();
        ModelSettings s;
        if (rdr.Read())
        {
            s = new ModelSettings
            {
                ModelId = rdr.GetInt64(0),
                OmitEmptyRows = rdr.GetInt32(1) != 0,
                OmitEmptyColumns = rdr.GetInt32(2) != 0,
                MemberDisplay = rdr.GetInt32(3),
                PreserveFormulas = rdr.GetInt32(4) != 0
            };
        }
        else
        {
            s = new ModelSettings { ModelId = modelId };
        }
        _settingsCache[modelId] = s;
        return s;
    }

    public void SaveSettings(ModelSettings s)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO ModelSettings (ModelId, OmitEmptyRows, OmitEmptyColumns, MemberDisplay, PreserveFormulas)
VALUES ($m, $r, $c, $d, $p)
ON CONFLICT(ModelId) DO UPDATE SET
    OmitEmptyRows = excluded.OmitEmptyRows,
    OmitEmptyColumns = excluded.OmitEmptyColumns,
    MemberDisplay = excluded.MemberDisplay,
    PreserveFormulas = excluded.PreserveFormulas";
        cmd.Parameters.AddWithValue("$m", s.ModelId);
        cmd.Parameters.AddWithValue("$r", s.OmitEmptyRows ? 1 : 0);
        cmd.Parameters.AddWithValue("$c", s.OmitEmptyColumns ? 1 : 0);
        cmd.Parameters.AddWithValue("$d", s.MemberDisplay);
        cmd.Parameters.AddWithValue("$p", s.PreserveFormulas ? 1 : 0);
        cmd.ExecuteNonQuery();
        InvalidateSettingsCache(s.ModelId);
    }

    #endregion

    #region ModelViews (persisted OLAP layout)

    public void SaveModelView(long modelId, string viewPayload)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"
INSERT INTO ModelViews (ModelId, ViewPayload, UpdatedUtc)
VALUES ($m, $p, $u)
ON CONFLICT(ModelId) DO UPDATE SET
    ViewPayload = excluded.ViewPayload,
    UpdatedUtc = excluded.UpdatedUtc";
        cmd.Parameters.AddWithValue("$m", modelId);
        cmd.Parameters.AddWithValue("$p", viewPayload ?? "");
        cmd.Parameters.AddWithValue("$u", DateTime.UtcNow.ToString("o"));
        cmd.ExecuteNonQuery();
    }

    public string? LoadModelView(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT ViewPayload FROM ModelViews WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        var result = cmd.ExecuteScalar()?.ToString();
        return string.IsNullOrWhiteSpace(result) ? null : result;
    }

    public void DeleteModelView(long modelId)
    {
        var conn = GetConnection();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM ModelViews WHERE ModelId = $m";
        cmd.Parameters.AddWithValue("$m", modelId);
        cmd.ExecuteNonQuery();
    }

    #endregion

    public void Dispose()
    {
        _conn?.Dispose();
        _conn = null;
    }
}
