# DataStreamEngine — Real-Time Windows-to-Excel Streaming

A production-grade, low-latency C# engine that streams jewellery billing data (Customers, Invoices, Loans) to CSV files with zero corruption — designed for live Excel dashboard consumption.

## Architecture

```
┌─────────────────────────────────────────────────┐
│              DataStreamService                   │
│         (Orchestrator + Data Generator)          │
├─────────────┬───────────────┬───────────────────┤
│ CsvDataWriter│ FileLockManager│ FileStreamLogger │
│ (Atomic CSV) │ (SemaphoreSlim)│ (Async Buffered) │
├─────────────┴───────────────┴───────────────────┤
│                  CSV Files                       │
│    customers.csv │ invoices.csv │ loans.csv      │
├─────────────────────────────────────────────────┤
│          Excel Dashboard (VBA Auto-Refresh)      │
└─────────────────────────────────────────────────┘
```

## Key Features

| Feature | Implementation |
|---|---|
| **Async I/O** | `FileStream` with 65KB buffering, fully async pipeline |
| **Atomic Writes** | Write-to-`.tmp` → NTFS rename (zero corruption) |
| **Concurrency** | Per-file `SemaphoreSlim` with configurable timeout |
| **Scalability** | 10,000+ entries per file, 5K row flush batching |
| **Logging** | Async buffered file logger with severity filtering |
| **Graceful Shutdown** | `CancellationToken` + `Ctrl+C` handler |
| **Excel Integration** | VBA `Application.OnTime` auto-refresh macro |

## Quick Start

### Prerequisites
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0)
- Microsoft Excel (for dashboard)

### Build & Run

```powershell
cd DataStreamEngine
dotnet build --configuration Release
dotnet run
```

The engine will:
1. Generate 10,000 customers, 10,000 invoices, and 5,000 loans
2. Write them atomically to `output/*.csv`
3. Begin streaming incremental updates every 3 seconds
4. Press **Ctrl+C** to stop gracefully

### Excel Dashboard
See [ExcelSetup.md](Excel/ExcelSetup.md) for the 3-minute dashboard setup guide.

## Configuration (`appsettings.json`)

| Setting | Default | Description |
|---|---|---|
| `OutputDirectory` | `output` | CSV output folder |
| `LogDirectory` | `logs` | Log file folder |
| `RefreshIntervalMs` | `3000` | Streaming cycle interval (ms) |
| `TotalCustomers` | `10000` | Initial customer count |
| `TotalInvoices` | `10000` | Initial invoice count |
| `TotalLoans` | `5000` | Initial loan count |
| `LockTimeoutMs` | `5000` | File lock acquisition timeout |
| `AtomicWriteEnabled` | `true` | Use atomic temp→rename writes |

## Project Structure

```
DataStreamEngine/
├── Core/
│   ├── Models/BillingModels.cs         # Customer, Invoice, Loan records
│   ├── Interfaces/Interfaces.cs        # IDataWriter, IStreamLogger, IConcurrencyGuard
│   └── Exceptions/DataStreamExceptions.cs
├── Infrastructure/
│   ├── Writers/CsvDataWriter.cs        # Atomic CSV writer (65KB buffer)
│   ├── Concurrency/FileLockManager.cs  # Per-file SemaphoreSlim locks
│   ├── Logging/FileStreamLogger.cs     # Async buffered logger
│   └── Configuration/StreamConfiguration.cs
├── Application/
│   └── Services/DataStreamService.cs   # Orchestrator + data generators
├── Excel/
│   └── ExcelSetup.md                   # Dashboard setup guide
├── Program.cs                          # Composition root
├── appsettings.json                    # Configuration
└── DataStreamEngine.csproj             # .NET 8 project
```

## License

Internal use — Billing Application Suite.
