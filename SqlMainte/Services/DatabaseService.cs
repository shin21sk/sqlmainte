using Microsoft.Data.SqlClient;
using SqlMainte.Models;

namespace SqlMainte.Services;

public class DatabaseService(string connectionString)
{
    /// <summary>テーブル全件取得。バイナリ列はカンマ区切り文字列に変換して返す。</summary>
    public List<Dictionary<string, object?>> FetchAll(string tableName, List<ColumnInfo> columns)
    {
        var sql = $"SELECT {string.Join(", ", columns.Select(c => $"[{c.ColumnName}]"))} FROM [{tableName}]";
        var rows = new List<Dictionary<string, object?>>();

        using var conn = new SqlConnection(connectionString);
        conn.Open();
        using var cmd = new SqlCommand(sql, conn);
        using var reader = cmd.ExecuteReader();

        while (reader.Read())
        {
            var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (var col in columns)
            {
                var value = reader.IsDBNull(reader.GetOrdinal(col.ColumnName))
                    ? null
                    : reader.GetValue(reader.GetOrdinal(col.ColumnName));

                row[col.ColumnName] = col.IsBinary
                    ? BinaryColumnSerializer.ToDisplayString(value as byte[])
                    : value;
            }
            rows.Add(row);
        }

        return rows;
    }

    /// <summary>変更をトランザクションで一括保存する。</summary>
    public void SaveChanges(
        string tableName,
        List<ColumnInfo> columns,
        List<Dictionary<string, object?>> toInsert,
        List<Dictionary<string, object?>> toUpdate,
        List<Dictionary<string, object?>> toDelete)
    {
        using var conn = new SqlConnection(connectionString);
        conn.Open();
        using var tx = conn.BeginTransaction();

        try
        {
            foreach (var row in toDelete)
                ExecuteDelete(conn, tx, tableName, columns, row);

            foreach (var row in toInsert)
                ExecuteInsert(conn, tx, tableName, columns, row);

            foreach (var row in toUpdate)
                ExecuteUpdate(conn, tx, tableName, columns, row);

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    private static void ExecuteInsert(
        SqlConnection conn, SqlTransaction tx,
        string tableName, List<ColumnInfo> columns,
        Dictionary<string, object?> row)
    {
        // IDENTITYかつ値が空の列は除外
        var insertCols = columns
            .Where(c => !(c.IsIdentity && IsNullOrEmpty(row, c.ColumnName)))
            .ToList();

        var colList = string.Join(", ", insertCols.Select(c => $"[{c.ColumnName}]"));
        var paramList = string.Join(", ", insertCols.Select(c => $"@{c.ColumnName}"));
        var sql = $"INSERT INTO [{tableName}] ({colList}) VALUES ({paramList})";

        using var cmd = new SqlCommand(sql, conn, tx);
        foreach (var col in insertCols)
            cmd.Parameters.Add(BuildParameter(col, row));

        cmd.ExecuteNonQuery();
    }

    private static void ExecuteUpdate(
        SqlConnection conn, SqlTransaction tx,
        string tableName, List<ColumnInfo> columns,
        Dictionary<string, object?> row)
    {
        var pkCols = columns.Where(c => c.IsPrimaryKey).ToList();
        var updateCols = columns.Where(c => !c.IsPrimaryKey).ToList();

        var setClauses = string.Join(", ", updateCols.Select(c => $"[{c.ColumnName}] = @{c.ColumnName}"));
        var whereClauses = string.Join(" AND ", pkCols.Select(c => $"[{c.ColumnName}] = @pk_{c.ColumnName}"));
        var sql = $"UPDATE [{tableName}] SET {setClauses} WHERE {whereClauses}";

        using var cmd = new SqlCommand(sql, conn, tx);
        foreach (var col in updateCols)
            cmd.Parameters.Add(BuildParameter(col, row));
        foreach (var pk in pkCols)
        {
            var p = BuildParameter(pk, row);
            p.ParameterName = $"@pk_{pk.ColumnName}";
            cmd.Parameters.Add(p);
        }

        cmd.ExecuteNonQuery();
    }

    private static void ExecuteDelete(
        SqlConnection conn, SqlTransaction tx,
        string tableName, List<ColumnInfo> columns,
        Dictionary<string, object?> row)
    {
        var pkCols = columns.Where(c => c.IsPrimaryKey).ToList();
        var whereClauses = string.Join(" AND ", pkCols.Select(c => $"[{c.ColumnName}] = @{c.ColumnName}"));
        var sql = $"DELETE FROM [{tableName}] WHERE {whereClauses}";

        using var cmd = new SqlCommand(sql, conn, tx);
        foreach (var pk in pkCols)
            cmd.Parameters.Add(BuildParameter(pk, row));

        cmd.ExecuteNonQuery();
    }

    private static SqlParameter BuildParameter(ColumnInfo col, Dictionary<string, object?> row)
    {
        var paramName = $"@{col.ColumnName}";
        var rawValue = row.TryGetValue(col.ColumnName, out var v) ? v : null;

        if (col.IsBinary)
        {
            var bytes = BinaryColumnSerializer.FromDisplayString(rawValue?.ToString());
            return new SqlParameter(paramName, bytes);
        }

        if (rawValue is null || (rawValue is string s && s == string.Empty && col.IsNullable))
            return new SqlParameter(paramName, DBNull.Value);

        return new SqlParameter(paramName, rawValue);
    }

    private static bool IsNullOrEmpty(Dictionary<string, object?> row, string colName)
        => !row.TryGetValue(colName, out var v) || v is null || v.ToString() == string.Empty;
}
