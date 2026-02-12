using BillingApp.Models;
using ClosedXML.Excel;
using System.IO;

namespace BillingApp.Services;

/// <summary>
/// Writes billing data to a native .xlsx Excel file using ClosedXML.
/// Event-driven: called only when the user submits a form.
/// Thread-safe via SemaphoreSlim.
/// </summary>
public sealed class ExcelWriterService : IDisposable
{
    private readonly string _excelPath;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private bool _disposed;

    public ExcelWriterService(string excelPath)
    {
        _excelPath = Path.GetFullPath(excelPath);
        EnsureWorkbookExists();
    }

    /// <summary>Create the workbook with headers if it doesn't exist.</summary>
    private void EnsureWorkbookExists()
    {
        if (File.Exists(_excelPath)) return;

        var dir = Path.GetDirectoryName(_excelPath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        using var wb = new XLWorkbook();

        // Customers sheet
        var cws = wb.AddWorksheet("Customers");
        var cHeaders = new[] { "ID", "Name", "Phone", "Address", "GSTIN", "Total Purchases", "Active Loans", "Loyalty Points", "Join Date" };
        for (int i = 0; i < cHeaders.Length; i++)
            cws.Cell(1, i + 1).Value = cHeaders[i];
        StyleHeaderRow(cws, cHeaders.Length);

        // Invoices sheet
        var iws = wb.AddWorksheet("Invoices");
        var iHeaders = new[] { "ID", "Customer ID", "Date", "Bill Type", "Item", "Metal", "Weight(g)", "Purity", "Rate/g", "Making", "Discount", "Sub Total", "CGST%", "SGST%", "IGST%", "GST Amount", "Total", "Status" };
        for (int i = 0; i < iHeaders.Length; i++)
            iws.Cell(1, i + 1).Value = iHeaders[i];
        StyleHeaderRow(iws, iHeaders.Length);

        // Loans sheet
        var lws = wb.AddWorksheet("Loans");
        var lHeaders = new[] { "ID", "Customer Name", "Phone", "Gov ID", "Metal", "Product", "Weight(g)", "Purity", "Principal", "Interest%", "Start Date", "Due Date", "Repaid", "Status" };
        for (int i = 0; i < lHeaders.Length; i++)
            lws.Cell(1, i + 1).Value = lHeaders[i];
        StyleHeaderRow(lws, lHeaders.Length);

        wb.SaveAs(_excelPath);
    }

    private static void StyleHeaderRow(IXLWorksheet ws, int colCount)
    {
        var headerRange = ws.Range(1, 1, 1, colCount);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Font.FontColor = XLColor.White;
        headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(40, 40, 60);
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        headerRange.Style.Border.BottomBorder = XLBorderStyleValues.Thick;
        headerRange.Style.Border.BottomBorderColor = XLColor.FromArgb(100, 100, 255);

        for (int i = 1; i <= colCount; i++)
            ws.Column(i).Width = 18;
    }

    // ─── Write Methods (Event-Driven) ──────────────────────────────────

    public async Task<int> WriteCustomerAsync(Customer c)
    {
        await _lock.WaitAsync();
        try
        {
            using var wb = new XLWorkbook(_excelPath);
            var ws = wb.Worksheet("Customers");
            int row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

            ws.Cell(row, 1).Value = c.Id;
            ws.Cell(row, 2).Value = c.Name;
            ws.Cell(row, 3).Value = c.Phone;
            ws.Cell(row, 4).Value = c.Address;
            ws.Cell(row, 5).Value = c.Gstin;
            ws.Cell(row, 6).Value = c.TotalPurchases;
            ws.Cell(row, 7).Value = c.ActiveLoans;
            ws.Cell(row, 8).Value = c.LoyaltyPoints;
            ws.Cell(row, 9).Value = c.JoinDate;

            StyleDataRow(ws, row, 9);
            wb.Save();
            return row - 1; // return row count
        }
        finally { _lock.Release(); }
    }

    public async Task<int> WriteInvoiceAsync(Invoice inv)
    {
        await _lock.WaitAsync();
        try
        {
            using var wb = new XLWorkbook(_excelPath);
            var ws = wb.Worksheet("Invoices");
            int row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

            ws.Cell(row, 1).Value = inv.Id;
            ws.Cell(row, 2).Value = inv.CustomerId;
            ws.Cell(row, 3).Value = inv.Date;
            ws.Cell(row, 4).Value = inv.BillType;
            ws.Cell(row, 5).Value = inv.ItemDescription;
            ws.Cell(row, 6).Value = inv.Metal;
            ws.Cell(row, 7).Value = inv.Weight;
            ws.Cell(row, 8).Value = inv.Purity;
            ws.Cell(row, 9).Value = inv.RatePerGram;
            ws.Cell(row, 10).Value = inv.MakingCharges;
            ws.Cell(row, 11).Value = inv.Discount;
            ws.Cell(row, 12).Value = inv.SubTotal;
            ws.Cell(row, 13).Value = inv.CgstRate;
            ws.Cell(row, 14).Value = inv.SgstRate;
            ws.Cell(row, 15).Value = inv.IgstRate;
            ws.Cell(row, 16).Value = inv.GstAmount;
            ws.Cell(row, 17).Value = inv.TotalAmount;
            ws.Cell(row, 18).Value = inv.Status;

            StyleDataRow(ws, row, 18);
            wb.Save();
            return row - 1;
        }
        finally { _lock.Release(); }
    }

    public async Task<int> WriteLoanAsync(Loan loan)
    {
        await _lock.WaitAsync();
        try
        {
            using var wb = new XLWorkbook(_excelPath);
            var ws = wb.Worksheet("Loans");
            int row = ws.LastRowUsed()?.RowNumber() + 1 ?? 2;

            ws.Cell(row, 1).Value = loan.Id;
            ws.Cell(row, 2).Value = loan.CustomerName;
            ws.Cell(row, 3).Value = loan.CustomerPhone;
            ws.Cell(row, 4).Value = loan.GovId;
            ws.Cell(row, 5).Value = loan.MetalType;
            ws.Cell(row, 6).Value = loan.ProductDescription;
            ws.Cell(row, 7).Value = loan.Weight;
            ws.Cell(row, 8).Value = loan.Purity;
            ws.Cell(row, 9).Value = loan.PrincipalAmount;
            ws.Cell(row, 10).Value = loan.InterestRate;
            ws.Cell(row, 11).Value = loan.StartDate;
            ws.Cell(row, 12).Value = loan.DueDate;
            ws.Cell(row, 13).Value = loan.TotalRepaid;
            ws.Cell(row, 14).Value = loan.Status;

            StyleDataRow(ws, row, 14);
            wb.Save();
            return row - 1;
        }
        finally { _lock.Release(); }
    }

    // ─── Read Methods (Load on Startup) ────────────────────────────────

    public List<Customer> LoadCustomers()
    {
        if (!File.Exists(_excelPath)) return new();
        using var wb = new XLWorkbook(_excelPath);
        var ws = wb.Worksheet("Customers");
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        var list = new List<Customer>();

        for (int r = 2; r <= lastRow; r++)
        {
            list.Add(new Customer
            {
                Id = ws.Cell(r, 1).GetString(),
                Name = ws.Cell(r, 2).GetString(),
                Phone = ws.Cell(r, 3).GetString(),
                Address = ws.Cell(r, 4).GetString(),
                Gstin = ws.Cell(r, 5).GetString(),
                TotalPurchases = GetDecimal(ws, r, 6),
                ActiveLoans = (int)GetDecimal(ws, r, 7),
                LoyaltyPoints = (int)GetDecimal(ws, r, 8),
                JoinDate = ws.Cell(r, 9).GetString()
            });
        }
        return list;
    }

    public List<Invoice> LoadInvoices()
    {
        if (!File.Exists(_excelPath)) return new();
        using var wb = new XLWorkbook(_excelPath);
        var ws = wb.Worksheet("Invoices");
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        var list = new List<Invoice>();

        for (int r = 2; r <= lastRow; r++)
        {
            list.Add(new Invoice
            {
                Id = ws.Cell(r, 1).GetString(),
                CustomerId = ws.Cell(r, 2).GetString(),
                Date = ws.Cell(r, 3).GetString(),
                BillType = ws.Cell(r, 4).GetString(),
                ItemDescription = ws.Cell(r, 5).GetString(),
                Metal = ws.Cell(r, 6).GetString(),
                Weight = GetDecimal(ws, r, 7),
                Purity = ws.Cell(r, 8).GetString(),
                RatePerGram = GetDecimal(ws, r, 9),
                MakingCharges = GetDecimal(ws, r, 10),
                Discount = GetDecimal(ws, r, 11),
                SubTotal = GetDecimal(ws, r, 12),
                CgstRate = GetDecimal(ws, r, 13),
                SgstRate = GetDecimal(ws, r, 14),
                IgstRate = GetDecimal(ws, r, 15),
                GstAmount = GetDecimal(ws, r, 16),
                TotalAmount = GetDecimal(ws, r, 17),
                Status = ws.Cell(r, 18).GetString()
            });
        }
        return list;
    }

    public List<Loan> LoadLoans()
    {
        if (!File.Exists(_excelPath)) return new();
        using var wb = new XLWorkbook(_excelPath);
        var ws = wb.Worksheet("Loans");
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        var list = new List<Loan>();

        for (int r = 2; r <= lastRow; r++)
        {
            list.Add(new Loan
            {
                Id = ws.Cell(r, 1).GetString(),
                CustomerName = ws.Cell(r, 2).GetString(),
                CustomerPhone = ws.Cell(r, 3).GetString(),
                GovId = ws.Cell(r, 4).GetString(),
                MetalType = ws.Cell(r, 5).GetString(),
                ProductDescription = ws.Cell(r, 6).GetString(),
                Weight = GetDecimal(ws, r, 7),
                Purity = ws.Cell(r, 8).GetString(),
                PrincipalAmount = GetDecimal(ws, r, 9),
                InterestRate = GetDecimal(ws, r, 10),
                StartDate = ws.Cell(r, 11).GetString(),
                DueDate = ws.Cell(r, 12).GetString(),
                TotalRepaid = GetDecimal(ws, r, 13),
                Status = ws.Cell(r, 14).GetString()
            });
        }
        return list;
    }

    /// <summary>Get the next auto-incremented ID for a sheet.</summary>
    public string GetNextId(string sheetName, string prefix)
    {
        if (!File.Exists(_excelPath)) return $"{prefix}001";
        using var wb = new XLWorkbook(_excelPath);
        var ws = wb.Worksheet(sheetName);
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 1;
        return $"{prefix}{(lastRow):D3}";
    }

    private static decimal GetDecimal(IXLWorksheet ws, int row, int col)
    {
        var cell = ws.Cell(row, col);
        if (cell.IsEmpty()) return 0;
        try { return (decimal)cell.GetDouble(); }
        catch { return 0; }
    }

    private static void StyleDataRow(IXLWorksheet ws, int row, int colCount)
    {
        var range = ws.Range(row, 1, row, colCount);
        range.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        range.Style.Border.BottomBorderColor = XLColor.FromArgb(60, 60, 80);
        if (row % 2 == 0)
            range.Style.Fill.BackgroundColor = XLColor.FromArgb(30, 30, 45);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _lock.Dispose();
    }
}
