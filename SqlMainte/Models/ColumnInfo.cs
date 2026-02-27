namespace SqlMainte.Models;

public class ColumnInfo
{
    public string ColumnName { get; set; } = string.Empty;
    public string DataType { get; set; } = string.Empty;
    public bool IsNullable { get; set; }
    public int? MaxLength { get; set; }
    public bool IsIdentity { get; set; }
    public bool IsPrimaryKey { get; set; }

    /// <summary>varbinary / binary 列かどうか</summary>
    public bool IsBinary => DataType is "varbinary" or "binary" or "image";
}
