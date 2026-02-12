namespace DataStreamEngine.Core.Exceptions;

/// <summary>Base exception for all data stream errors.</summary>
public class DataStreamException : Exception
{
    public DataStreamException(string message) : base(message) { }
    public DataStreamException(string message, Exception inner) : base(message, inner) { }
}

/// <summary>Thrown when a file lock cannot be acquired within the configured timeout.</summary>
public class FileLockTimeoutException : DataStreamException
{
    public string ResourceName { get; }
    public TimeSpan Timeout { get; }

    public FileLockTimeoutException(string resourceName, TimeSpan timeout)
        : base($"Failed to acquire lock on '{resourceName}' within {timeout.TotalMilliseconds}ms")
    {
        ResourceName = resourceName;
        Timeout = timeout;
    }
}

/// <summary>Thrown when an atomic write operation fails mid-flight.</summary>
public class AtomicWriteException : DataStreamException
{
    public string TargetPath { get; }

    public AtomicWriteException(string targetPath, Exception inner)
        : base($"Atomic write failed for '{targetPath}': {inner.Message}", inner)
    {
        TargetPath = targetPath;
    }
}
