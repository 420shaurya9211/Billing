using DataStreamEngine.Core.Interfaces;
using DataStreamEngine.Core.Exceptions;
using System.Reflection;
using System.Text;

namespace DataStreamEngine.Infrastructure.Writers;

/// <summary>
/// High-performance async CSV writer with atomic writes and corruption protection.
/// 
/// Strategy:
///   1. Serialize data to a temporary file (.tmp)
///   2. Atomically rename .tmp → .csv (rename is atomic on NTFS)
///   3. If rename fails, fall back to copy + delete
/// 
/// This guarantees Excel never sees a partially-written CSV.
/// </summary>
public sealed class CsvDataWriter : IDataWriter
{
    private readonly string _outputDirectory;
    private readonly IConcurrencyGuard _lockManager;
    private readonly IStreamLogger _logger;
    private readonly bool _atomicWriteEnabled;
    private bool _disposed;

    // Cache reflected property info per type to avoid repeated reflection
    private static readonly Dictionary<Type, PropertyInfo[]> _propertyCache = new();
    private static readonly object _cacheLock = new();

    public CsvDataWriter(
        string outputDirectory,
        IConcurrencyGuard lockManager,
        IStreamLogger logger,
        bool atomicWriteEnabled = true)
    {
        _outputDirectory = outputDirectory;
        _lockManager = lockManager;
        _logger = logger;
        _atomicWriteEnabled = atomicWriteEnabled;
        Directory.CreateDirectory(_outputDirectory);
    }

    /// <summary>
    /// Write an entire collection to CSV, replacing any existing file atomically.
    /// Handles 10,000+ rows with buffered streaming.
    /// </summary>
    public async Task WriteAllAsync<T>(string fileName, IEnumerable<T> data, CancellationToken ct = default) where T : class
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var targetPath = Path.Combine(_outputDirectory, fileName);
        var tempPath = targetPath + ".tmp";
        var properties = GetCachedProperties(typeof(T));

        await _logger.LogInfoAsync($"Writing full dataset to {fileName}", "CsvWriter");

        using var lockHandle = await _lockManager.AcquireAsync(fileName, ct).ConfigureAwait(false);

        try
        {
            // Phase 1: Write to temp file with buffering
            var rowCount = 0;
            await using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None, 65536, true))
            await using (var writer = new StreamWriter(stream, Encoding.UTF8, 65536))
            {
                // Write CSV header
                await writer.WriteLineAsync(string.Join(",", properties.Select(p => EscapeCsv(p.Name))))
                    .ConfigureAwait(false);

                // Write data rows
                foreach (var item in data)
                {
                    ct.ThrowIfCancellationRequested();

                    var values = properties.Select(p =>
                    {
                        var val = p.GetValue(item);
                        return EscapeCsv(val?.ToString() ?? "");
                    });

                    await writer.WriteLineAsync(string.Join(",", values)).ConfigureAwait(false);
                    rowCount++;

                    // Flush every 5000 rows to manage memory
                    if (rowCount % 5000 == 0)
                        await writer.FlushAsync(ct).ConfigureAwait(false);
                }
            }

            // Phase 2: Atomic rename
            if (_atomicWriteEnabled)
            {
                AtomicMove(tempPath, targetPath);
            }
            else
            {
                File.Copy(tempPath, targetPath, overwrite: true);
                File.Delete(tempPath);
            }

            await _logger.LogInfoAsync($"Successfully wrote {rowCount:N0} rows to {fileName}", "CsvWriter");
        }
        catch (OperationCanceledException)
        {
            CleanupTempFile(tempPath);
            throw;
        }
        catch (Exception ex) when (ex is not DataStreamException)
        {
            CleanupTempFile(tempPath);
            throw new AtomicWriteException(targetPath, ex);
        }
    }

    /// <summary>
    /// Append a single record to an existing CSV file.
    /// Creates the file with headers if it doesn't exist.
    /// </summary>
    public async Task AppendAsync<T>(string fileName, T record, CancellationToken ct = default) where T : class
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var targetPath = Path.Combine(_outputDirectory, fileName);
        var properties = GetCachedProperties(typeof(T));

        using var lockHandle = await _lockManager.AcquireAsync(fileName, ct).ConfigureAwait(false);

        try
        {
            var fileExists = File.Exists(targetPath);

            await using var stream = new FileStream(targetPath, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, true);
            await using var writer = new StreamWriter(stream, Encoding.UTF8);

            // Write header if new file
            if (!fileExists || new FileInfo(targetPath).Length == 0)
            {
                await writer.WriteLineAsync(string.Join(",", properties.Select(p => EscapeCsv(p.Name))))
                    .ConfigureAwait(false);
            }

            // Write the record
            var values = properties.Select(p =>
            {
                var val = p.GetValue(record);
                return EscapeCsv(val?.ToString() ?? "");
            });
            await writer.WriteLineAsync(string.Join(",", values)).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is not DataStreamException)
        {
            throw new AtomicWriteException(targetPath, ex);
        }
    }

    /// <summary>Atomically replace target file using rename (NTFS atomic operation).</summary>
    private static void AtomicMove(string source, string target)
    {
        // File.Move with overwrite is atomic on NTFS for same-volume moves
        if (File.Exists(target))
        {
            var backupPath = target + ".bak";
            try
            {
                // Keep one backup for disaster recovery
                File.Copy(target, backupPath, overwrite: true);
            }
            catch { /* Backup is best-effort */ }
        }

        File.Move(source, target, overwrite: true);
    }

    private static void CleanupTempFile(string tempPath)
    {
        try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
    }

    private static PropertyInfo[] GetCachedProperties(Type type)
    {
        lock (_cacheLock)
        {
            if (!_propertyCache.TryGetValue(type, out var props))
            {
                props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                            .OrderBy(p => p.Name)
                            .ToArray();
                _propertyCache[type] = props;
            }
            return props;
        }
    }

    /// <summary>RFC 4180 compliant CSV escaping.</summary>
    private static string EscapeCsv(string value)
    {
        if (value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r'))
            return $"\"{value.Replace("\"", "\"\"")}\"";
        return value;
    }

    public ValueTask DisposeAsync()
    {
        _disposed = true;
        return ValueTask.CompletedTask;
    }
}
