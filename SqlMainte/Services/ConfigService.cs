using Microsoft.Extensions.Configuration;
using SqlMainte.Models;

namespace SqlMainte.Services;

public static class ConfigService
{
    private static AppSettings? _cache;

    public static AppSettings Load()
    {
        if (_cache is not null) return _cache;

        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
            .Build();

        _cache = new AppSettings
        {
            ConnectionString = config["ConnectionString"] ?? string.Empty,
            Tables = config.GetSection("Tables").Get<List<TableConfig>>() ?? []
        };

        return _cache;
    }
}
