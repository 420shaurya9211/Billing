namespace DataStreamEngine.Core.Interfaces;

/// <summary>
/// Async data writer that serializes collections to a persistent store.
/// </summary>
public interface IDataWriter : IAsyncDisposable
{
    /// <summary>Write an entire collection, replacing any existing data.</summary>
    Task WriteAllAsync<T>(string fileName, IEnumerable<T> data, CancellationToken ct = default) where T : class;

    /// <summary>Append a single record to the end of an existing file.</summary>
    Task AppendAsync<T>(string fileName, T record, CancellationToken ct = default) where T : class;
}

/// <summary>
/// Structured async logger with severity levels.
/// </summary>
public interface IStreamLogger : IAsyncDisposable
{
    Task LogInfoAsync(string message, string? context = null);
    Task LogWarningAsync(string message, string? context = null);
    Task LogErrorAsync(string message, Exception? ex = null, string? context = null);
}

/// <summary>
/// Manages concurrency guards for named resources (e.g. per-file locks).
/// </summary>
public interface IConcurrencyGuard : IDisposable
{
    /// <summary>Acquire a named lock. Returns an IDisposable that releases on dispose.</summary>
    Task<IDisposable> AcquireAsync(string resourceName, CancellationToken ct = default);
}
