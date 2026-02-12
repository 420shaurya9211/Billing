using Microsoft.Extensions.Configuration;

namespace DataStreamEngine.Infrastructure.Configuration;

/// <summary>
/// Strongly-typed configuration loaded from appsettings.json with validated defaults.
/// </summary>
public sealed class StreamConfiguration
{
    public string OutputDirectory { get; set; } = "output";
    public string LogDirectory { get; set; } = "logs";
    public int RefreshIntervalMs { get; set; } = 3000;
    public int MaxEntriesPerCycle { get; set; } = 500;
    public int TotalCustomers { get; set; } = 10000;
    public int TotalInvoices { get; set; } = 10000;
    public int TotalLoans { get; set; } = 5000;
    public int LockTimeoutMs { get; set; } = 5000;
    public string LogLevel { get; set; } = "Information";
    public bool EnableConsoleOutput { get; set; } = true;
    public bool AtomicWriteEnabled { get; set; } = true;

    /// <summary>Load and validate configuration from appsettings.json.</summary>
    public static StreamConfiguration Load()
    {
        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();

        var settings = new StreamConfiguration();
        config.GetSection("DataStream").Bind(settings);
        settings.Validate();
        return settings;
    }

    private void Validate()
    {
        if (RefreshIntervalMs < 100)
            throw new ArgumentException("RefreshIntervalMs must be >= 100ms");
        if (TotalCustomers < 1)
            throw new ArgumentException("TotalCustomers must be >= 1");
        if (LockTimeoutMs < 500)
            throw new ArgumentException("LockTimeoutMs must be >= 500ms");
    }
}
