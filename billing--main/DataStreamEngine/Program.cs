using DataStreamEngine.Application.Services;
using DataStreamEngine.Infrastructure.Concurrency;
using DataStreamEngine.Infrastructure.Configuration;
using DataStreamEngine.Infrastructure.Logging;
using DataStreamEngine.Infrastructure.Writers;

namespace DataStreamEngine;

/// <summary>
/// Composition root — wires all dependencies and runs the streaming engine.
/// Supports Ctrl+C graceful shutdown.
/// </summary>
public static class Program
{
    public static async Task Main(string[] args)
    {
        // ── Load Configuration ──
        StreamConfiguration config;
        try
        {
            config = StreamConfiguration.Load();
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[FATAL] Configuration error: {ex.Message}");
            Console.ResetColor();
            Environment.ExitCode = 1;
            return;
        }

        // ── Wire Dependencies ──
        await using var logger = new FileStreamLogger(
            config.LogDirectory,
            config.LogLevel,
            config.EnableConsoleOutput
        );

        using var lockManager = new FileLockManager(config.LockTimeoutMs);

        await using var csvWriter = new CsvDataWriter(
            config.OutputDirectory,
            lockManager,
            logger,
            config.AtomicWriteEnabled
        );

        await using var streamService = new DataStreamService(csvWriter, logger, config);

        // ── Graceful Shutdown ──
        using var cts = new CancellationTokenSource();

        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true; // Prevent immediate termination
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n[SHUTDOWN] Graceful shutdown initiated... flushing buffers.");
            Console.ResetColor();
            cts.Cancel();
        };

        AppDomain.CurrentDomain.ProcessExit += (_, _) =>
        {
            if (!cts.IsCancellationRequested) cts.Cancel();
        };

        // ── Run Engine ──
        try
        {
            await streamService.RunAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            await logger.LogInfoAsync("Engine stopped gracefully.", "Shutdown");
        }
        catch (Exception ex)
        {
            await logger.LogErrorAsync("Engine terminated with error", ex, "Fatal");
            Environment.ExitCode = 1;
        }

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("\n[DONE] DataStreamEngine shutdown complete. CSV files preserved.");
        Console.ResetColor();
    }
}
