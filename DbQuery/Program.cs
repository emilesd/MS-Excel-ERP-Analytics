using System;
using Microsoft.Data.Sqlite;

SQLitePCL.Batteries.Init();
var conn = new SqliteConnection("Data Source=C:/Users/devadmin/AppData/Local/MyOlap/myolap.db");
conn.Open();

// Count duplicates before
var cmd = conn.CreateCommand();
cmd.CommandText = "SELECT COUNT(DISTINCT MemberKey), COUNT(*) FROM FactData WHERE ModelId = 31";
var rdr = cmd.ExecuteReader();
rdr.Read();
Console.WriteLine($"Before cleanup: {rdr.GetInt64(0)} distinct keys, {rdr.GetInt64(1)} total facts");
rdr.Close();

// Delete duplicates, keeping the latest Id per MemberKey
cmd = conn.CreateCommand();
cmd.CommandText = @"DELETE FROM FactData WHERE ModelId = 31 AND Id NOT IN (
    SELECT MAX(Id) FROM FactData WHERE ModelId = 31 GROUP BY MemberKey
)";
var deleted = cmd.ExecuteNonQuery();
Console.WriteLine($"Deleted {deleted} duplicate facts");

// Count after
cmd = conn.CreateCommand();
cmd.CommandText = "SELECT COUNT(DISTINCT MemberKey), COUNT(*) FROM FactData WHERE ModelId = 31";
rdr = cmd.ExecuteReader();
rdr.Read();
Console.WriteLine($"After cleanup: {rdr.GetInt64(0)} distinct keys, {rdr.GetInt64(1)} total facts");
rdr.Close();

conn.Close();
