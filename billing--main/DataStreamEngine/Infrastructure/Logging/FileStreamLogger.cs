using DataStreamEngine.Core.Interfaces;
using System.Collections.Concurrent;
using System.Text;

namespace DataStreamEngine.Infrastructure.Logging;

/// <summary>
/// Production-grade async file logger with rolling output, severity levels, and thread safety.
/// </summary>
public sealed class FileStreamLogger : IStreamLogger
{
    private readonly string _logDirectory;
    private readonly string _logLevel;
    private readonly bool _enableConsole;
    private readonly SemaphoreSlim _writeLock = new(1, 1);
    private readonly ConcurrentQueue<string> _buffer = new();
    private readonly Timer _flushTimer;
    private bool _disposed;

    private static readonly Dictionary<string, int> LevelPriority = new()
    {
        ["Debug"] = 0, ["Information"] = 1, ["Warning"] = 2, ["Error"] = 3
    };

    public FileStreamLogger(string logDirectory, string logLevel = "Information", bool enableConsole = true)
    {
        _logDirectory = logDirectory;
        _logLevel = logLevel;
        _enableConsole = enableConsole;
        Directory.CreateDirectory(_logDirectory);

        // Flush buffer every 500ms for low-latency log persistence
        _flushTimer = new Timer(_ => _ = FlushBufferAsync(), null, 500, 500);
    }

    public Task LogInfoAsync(string message, string? context = null) =>
        WriteEntryAsync("INF", message, context);

    public Task LogWarningAsync(string message, string? context = null) =>
        WriteEntryAsync("WRN", message, context);

    public Task LogErrorAsync(string message, Exception? ex = null, string? context = null)
    {
        var fullMessage = ex is not null ? $"{message} | Exception: {ex.GetType().Name}: {ex.Message}" : message;
        return WriteEntryAsync("ERR", fullMessage, context);
    }

    private Task WriteEntryAsync(string level, string message, string? context)
    {
        // Check log level filtering
        var levelMap = new Dictionary<string, string> { ["INF"] = "Information", ["WRN"] = "Warning", ["ERR"] = "Error" };
        if (levelMap.TryGetValue(level, out var levelName) &&
            LevelPriority.TryGetValue(levelName, out var priority) &&
            LevelPriority.TryGetValue(_logLevel, out var minPriority) &&
            priority < minPriority)
        {
            return Task.CompletedTask;
        }

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var threadId = Environment.CurrentManagedThreadId;
        var ctx = context is not null ? $" [{context}]" : "";
        var entry = $"[{timestamp}] [{level}] [T{threadId:D3}]{ctx} {message}";

        _buffer.Enqueue(entry);

        if (_enableConsole)
        {
            var originalColor = Console.ForegroundColor;
            Console.ForegroundColor = level switch
            {
                "ERR" => ConsoleColor.Red,
                "WRN" => ConsoleColor.Yellow,
                "INF" => ConsoleColor.Cyan,
                _ => ConsoleColor.Gray
            };
            Console.WriteLine(entry);
            Console.ForegroundColor = originalColor;
        }

        return Task.CompletedTask;
    }

    private async Task FlushBufferAsync()
    {
        if (_buffer.IsEmpty) return;

        await _writeLock.WaitAsync().ConfigureAwait(false);
        try
        {
            var logFile = Path.Combine(_logDirectory, $"stream_{DateTime.Now:yyyy-MM-dd}.log");
            var sb = new StringBuilder();

            while (_buffer.TryDequeue(out var entry))
                sb.AppendLine(entry);

            if (sb.Length > 0)
                await File.AppendAllTextAsync(logFile, sb.ToString()).ConfigureAwait(false);
        }
        catch
        {
            // Logger must never throw — swallow exceptions in production
        }
        finally
        {
            _writeLock.Release();
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed) return;
        _disposed = true;

        await _flushTimer.DisposeAsync().ConfigureAwait(false);
        await FlushBufferAsync().ConfigureAwait(false);
        _writeLock.Dispose();
    }
}
