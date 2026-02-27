using Microsoft.Data.SqlClient;
using SqlMainte.Models;

namespace SqlMainte.Services;

public class SchemaService(string connectionString)
{
    public List<ColumnInfo> GetColumns(string tableName, List<string> primaryKeys)
    {
        const string sql = """
            SELECT
                c.COLUMN_NAME,
                c.DATA_TYPE,
                c.IS_NULLABLE,
                c.CHARACTER_MAXIMUM_LENGTH,
                COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME), c.COLUMN_NAME, 'IsIdentity') AS IS_IDENTITY
            FROM INFORMATION_SCHEMA.COLUMNS c
            WHERE c.TABLE_NAME = @TableName
            ORDER BY c.ORDINAL_POSITION
            """;

        var columns = new List<ColumnInfo>();

        using var conn = new SqlConnection(connectionString);
        conn.Open();
        using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@TableName", tableName);

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var colName = reader.GetString(0);
            columns.Add(new ColumnInfo
            {
                ColumnName = colName,
                DataType = reader.GetString(1),
                IsNullable = reader.GetString(2) == "YES",
                MaxLength = reader.IsDBNull(3) ? null : reader.GetInt32(3),
                IsIdentity = reader.GetInt32(4) == 1,
                IsPrimaryKey = primaryKeys.Contains(colName, StringComparer.OrdinalIgnoreCase)
            });
        }

        return columns;
    }
}
