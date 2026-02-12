using DataStreamEngine.Core.Interfaces;
using DataStreamEngine.Core.Exceptions;
using System.Collections.Concurrent;

namespace DataStreamEngine.Infrastructure.Concurrency;

/// <summary>
/// Per-resource SemaphoreSlim manager with configurable timeout.
/// Prevents file corruption from concurrent writes via ordered lock acquisition.
/// </summary>
public sealed class FileLockManager : IConcurrencyGuard
{
    private readonly ConcurrentDictionary<string, SemaphoreSlim> _locks = new();
    private readonly TimeSpan _timeout;
    private bool _disposed;

    public FileLockManager(int timeoutMs = 5000)
    {
        _timeout = TimeSpan.FromMilliseconds(timeoutMs);
    }

    public async Task<IDisposable> AcquireAsync(string resourceName, CancellationToken ct = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);

        var semaphore = _locks.GetOrAdd(resourceName, _ => new SemaphoreSlim(1, 1));

        bool acquired = await semaphore.WaitAsync(_timeout, ct).ConfigureAwait(false);
        if (!acquired)
            throw new FileLockTimeoutException(resourceName, _timeout);

        return new LockRelease(semaphore);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        foreach (var kvp in _locks)
        {
            kvp.Value.Dispose();
        }
        _locks.Clear();
    }

    /// <summary>RAII-style lock releaser.</summary>
    private sealed class LockRelease : IDisposable
    {
        private readonly SemaphoreSlim _semaphore;
        private int _released;

        public LockRelease(SemaphoreSlim semaphore) => _semaphore = semaphore;

        public void Dispose()
        {
            if (Interlocked.CompareExchange(ref _released, 1, 0) == 0)
                _semaphore.Release();
        }
    }
}
