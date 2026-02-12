using DataStreamEngine.Core.Interfaces;
using DataStreamEngine.Core.Models;
using DataStreamEngine.Infrastructure.Configuration;

namespace DataStreamEngine.Application.Services;

/// <summary>
/// Orchestrates the data generation → CSV writing → logging pipeline.
/// Runs on a configurable periodic timer with graceful cancellation support.
/// </summary>
public sealed class DataStreamService : IAsyncDisposable
{
    private readonly IDataWriter _writer;
    private readonly IStreamLogger _logger;
    private readonly StreamConfiguration _config;
    private readonly Random _rng = new();

    // Realistic Indian jewellery shop data
    private static readonly string[] FirstNames = { "Rajesh", "Priya", "Amit", "Sunita", "Vikram", "Anjali", "Deepak", "Meena", "Suresh", "Kavita", "Rohit", "Neha", "Arun", "Pooja", "Ravi", "Sita", "Manoj", "Rekha", "Ashok", "Lata", "Gaurav", "Swati", "Sanjay", "Anita", "Nitin" };
    private static readonly string[] LastNames = { "Sharma", "Patel", "Verma", "Gupta", "Singh", "Jain", "Agarwal", "Mehta", "Shah", "Chopra", "Kumar", "Desai", "Modi", "Trivedi", "Reddy", "Nair", "Iyer", "Pillai", "Das", "Bose" };
    private static readonly string[] Cities = { "Mumbai", "Delhi", "Ahmedabad", "Jaipur", "Bangalore", "Chennai", "Kolkata", "Hyderabad", "Pune", "Lucknow", "Surat", "Nagpur", "Indore", "Varanasi", "Chandigarh" };
    private static readonly string[] Areas = { "Bandra West", "Karol Bagh", "Navrangpura", "Johari Bazaar", "Commercial St", "T Nagar", "New Market", "Charminar", "MG Road", "Hazratganj", "Ring Road", "Sadar", "Rajwada", "Godowlia", "Sector 17" };
    private static readonly string[] GoldItems = { "Gold Necklace 22K", "Gold Bangle 22K", "Gold Ring 22K", "Gold Chain 22K", "Gold Earrings 22K", "Gold Pendant 18K", "Gold Bracelet 22K", "Gold Anklet 22K", "Gold Nose Pin 22K", "Gold Mangalsutra 22K" };
    private static readonly string[] SilverItems = { "Silver Payal", "Silver Glass Set", "Silver Thali", "Silver Idol", "Silver Coin 999", "Silver Bangle", "Silver Ring", "Silver Chain" };
    private static readonly string[] LoanItems = { "Gold Chain (Man)", "Gold Necklace Set", "Gold Bangles (4 pcs)", "Gold Ring", "Gold Earrings", "Gold Bracelet" };
    private static readonly string[] Purities = { "24K", "22K", "18K", "14K" };

    public DataStreamService(IDataWriter writer, IStreamLogger logger, StreamConfiguration config)
    {
        _writer = writer;
        _logger = logger;
        _config = config;
    }

    /// <summary>
    /// Run the streaming engine: generates data in batches on a periodic interval.
    /// </summary>
    public async Task RunAsync(CancellationToken ct)
    {
        await _logger.LogInfoAsync("═══════════════════════════════════════════════════════", "Engine");
        await _logger.LogInfoAsync("  DataStreamEngine v1.0 — Real-Time Excel Streaming", "Engine");
        await _logger.LogInfoAsync("═══════════════════════════════════════════════════════", "Engine");
        await _logger.LogInfoAsync($"Config → Customers: {_config.TotalCustomers:N0} | Invoices: {_config.TotalInvoices:N0} | Loans: {_config.TotalLoans:N0}", "Engine");
        await _logger.LogInfoAsync($"Config → Refresh: {_config.RefreshIntervalMs}ms | Output: {_config.OutputDirectory}", "Engine");
        await _logger.LogInfoAsync("Press Ctrl+C to stop gracefully.", "Engine");
        await _logger.LogInfoAsync("", "Engine");

        // --- Phase 1: Generate and write full dataset ---
        await _logger.LogInfoAsync("▸ Phase 1: Generating full dataset...", "Engine");

        var customers = GenerateCustomers(_config.TotalCustomers);
        await _writer.WriteAllAsync("customers.csv", customers, ct);

        var invoices = GenerateInvoices(_config.TotalInvoices, customers);
        await _writer.WriteAllAsync("invoices.csv", invoices, ct);

        var loans = GenerateLoans(_config.TotalLoans, customers);
        await _writer.WriteAllAsync("loans.csv", loans, ct);

        await _logger.LogInfoAsync($"✓ Phase 1 complete — {customers.Count + invoices.Count + loans.Count:N0} total records written", "Engine");
        await _logger.LogInfoAsync("", "Engine");

        // --- Phase 2: Streaming incremental updates ---
        await _logger.LogInfoAsync("▸ Phase 2: Starting real-time incremental streaming...", "Engine");

        using var timer = new PeriodicTimer(TimeSpan.FromMilliseconds(_config.RefreshIntervalMs));
        var cycle = 0;

        while (await timer.WaitForNextTickAsync(ct).ConfigureAwait(false))
        {
            cycle++;
            try
            {
                var batchSize = _rng.Next(1, Math.Min(10, _config.MaxEntriesPerCycle));
                await _logger.LogInfoAsync($"Cycle {cycle}: Streaming {batchSize} new entries...", "Stream");

                // Randomly add new customers, invoices, or loans
                for (int i = 0; i < batchSize; i++)
                {
                    var type = _rng.Next(3);
                    switch (type)
                    {
                        case 0:
                            var newCustomer = GenerateCustomers(1).First();
                            await _writer.AppendAsync("customers.csv", newCustomer, ct);
                            break;
                        case 1:
                            var newInvoice = GenerateInvoices(1, customers).First();
                            await _writer.AppendAsync("invoices.csv", newInvoice, ct);
                            break;
                        case 2:
                            var newLoan = GenerateLoans(1, customers).First();
                            await _writer.AppendAsync("loans.csv", newLoan, ct);
                            break;
                    }
                }

                await _logger.LogInfoAsync($"Cycle {cycle}: ✓ Batch committed", "Stream");
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                await _logger.LogErrorAsync($"Cycle {cycle} failed", ex, "Stream");
            }
        }
    }

    // ─── Data Generators ──────────────────────────────────────────────────────

    private List<Customer> GenerateCustomers(int count)
    {
        var customers = new List<Customer>(count);
        for (int i = 0; i < count; i++)
        {
            var firstName = FirstNames[_rng.Next(FirstNames.Length)];
            var lastName = LastNames[_rng.Next(LastNames.Length)];
            var city = Cities[_rng.Next(Cities.Length)];
            var area = Areas[_rng.Next(Areas.Length)];
            var totalPurchases = _rng.Next(10000, 5000000);

            customers.Add(new Customer
            {
                Id = $"C{(i + 1):D5}",
                Name = $"{firstName} {lastName}",
                Phone = $"9{_rng.Next(100000000, 999999999)}",
                Address = $"{area}, {city}",
                Gstin = _rng.Next(4) == 0 ? $"{_rng.Next(10, 36)}{GenerateAlphaNum(10)}" : null,
                TotalPurchases = totalPurchases,
                ActiveLoans = _rng.Next(0, 4),
                LoyaltyPoints = totalPurchases / 1000,
                JoinDate = RandomDate(2020, 2025)
            });
        }
        return customers;
    }

    private List<Invoice> GenerateInvoices(int count, List<Customer> customers)
    {
        var invoices = new List<Invoice>(count);
        for (int i = 0; i < count; i++)
        {
            var customer = customers[_rng.Next(customers.Count)];
            var isGold = _rng.Next(3) > 0; // 66% gold
            var items = isGold ? GoldItems : SilverItems;
            var item = items[_rng.Next(items.Length)];
            var weight = Math.Round(_rng.NextDouble() * 50 + 1, 2);
            var rate = isGold ? _rng.Next(6500, 7800) : _rng.Next(75, 110);
            var purity = isGold ? Purities[_rng.Next(0, 3)] : "999";
            var making = _rng.Next(100, 500);
            var discount = _rng.Next(0, 3000);
            var subTotal = Math.Round((decimal)(weight * rate + making - discount), 2);
            var cgst = Math.Round(subTotal * 0.015m, 2);
            var sgst = cgst;
            var total = subTotal + cgst + sgst;

            invoices.Add(new Invoice
            {
                Id = $"INV-{(i + 1):D6}",
                CustomerId = customer.Id,
                Date = RandomDate(2024, 2025),
                BillType = _rng.Next(2) == 0 ? "PAKKA" : "KACHA",
                ItemDescription = item,
                Metal = isGold ? "GOLD" : "SILVER",
                Weight = (decimal)weight,
                Purity = purity,
                RatePerGram = rate,
                MakingCharges = making,
                Discount = discount,
                SubTotal = subTotal,
                CgstRate = 1.5m,
                SgstRate = 1.5m,
                IgstRate = 0,
                GstAmount = cgst + sgst,
                TotalAmount = total,
                Status = _rng.Next(5) == 0 ? "PENDING" : "PAID"
            });
        }
        return invoices;
    }

    private List<Loan> GenerateLoans(int count, List<Customer> customers)
    {
        var loans = new List<Loan>(count);
        var statuses = new[] { "ACTIVE", "ACTIVE", "ACTIVE", "CLOSED", "OVERDUE" };

        for (int i = 0; i < count; i++)
        {
            var customer = customers[_rng.Next(customers.Count)];
            var weight = Math.Round(_rng.NextDouble() * 30 + 2, 2);
            var principal = _rng.Next(10000, 500000);

            loans.Add(new Loan
            {
                Id = $"L-{(i + 1):D5}",
                CustomerName = customer.Name,
                CustomerPhone = customer.Phone,
                GovId = GenerateAlphaNum(10).ToUpper(),
                MetalType = "GOLD",
                ProductDescription = LoanItems[_rng.Next(LoanItems.Length)],
                Weight = (decimal)weight,
                Purity = "22K",
                PrincipalAmount = principal,
                InterestRate = Math.Round((decimal)(_rng.NextDouble() * 2 + 0.5), 2),
                StartDate = RandomDate(2023, 2025),
                DueDate = RandomDate(2025, 2026),
                TotalRepaid = _rng.Next(0, principal / 2),
                Status = statuses[_rng.Next(statuses.Length)]
            });
        }
        return loans;
    }

    private string RandomDate(int startYear, int endYear)
    {
        var start = new DateTime(startYear, 1, 1);
        var range = (new DateTime(endYear, 12, 31) - start).Days;
        return start.AddDays(_rng.Next(range)).ToString("yyyy-MM-dd");
    }

    private string GenerateAlphaNum(int length)
    {
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        return new string(Enumerable.Range(0, length).Select(_ => chars[_rng.Next(chars.Length)]).ToArray());
    }

    public async ValueTask DisposeAsync()
    {
        await _writer.DisposeAsync();
        await _logger.DisposeAsync();
    }
}
