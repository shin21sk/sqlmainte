namespace SqlMainte.Models;

public class AppSettings
{
    public string ConnectionString { get; set; } = string.Empty;
    public List<TableConfig> Tables { get; set; } = [];
}

public class TableConfig
{
    public string Name { get; set; } = string.Empty;
    public string DisplayName { get; set; } = string.Empty;
    public List<string> PrimaryKeys { get; set; } = [];
}
