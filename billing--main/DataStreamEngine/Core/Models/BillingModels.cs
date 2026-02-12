namespace DataStreamEngine.Core.Models;

/// <summary>
/// Represents a jewellery shop customer with purchase history and loyalty tracking.
/// </summary>
public sealed record Customer
{
    public required string Id { get; init; }
    public required string Name { get; init; }
    public required string Phone { get; init; }
    public required string Address { get; init; }
    public string? Gstin { get; init; }
    public decimal TotalPurchases { get; init; }
    public int ActiveLoans { get; init; }
    public int LoyaltyPoints { get; init; }
    public required string JoinDate { get; init; }
}

/// <summary>
/// A single line item on a jewellery invoice.
/// </summary>
public sealed record InvoiceItem
{
    public required string Id { get; init; }
    public required string Description { get; init; }
    public required string Metal { get; init; }   // GOLD | SILVER
    public decimal Weight { get; init; }
    public required string Purity { get; init; }
    public decimal RatePerGram { get; init; }
    public decimal MakingCharges { get; init; }
    public required string MakingChargeType { get; init; } // FIXED | PERCENTAGE
    public decimal Discount { get; init; }
    public decimal TaxableValue { get; init; }
    public string? HsnCode { get; init; }
}

/// <summary>
/// Jewellery invoice with GST breakdown.
/// </summary>
public sealed record Invoice
{
    public required string Id { get; init; }
    public required string CustomerId { get; init; }
    public required string Date { get; init; }
    public string? BillType { get; init; }   // PAKKA | KACHA
    public required string ItemDescription { get; init; }
    public required string Metal { get; init; }
    public decimal Weight { get; init; }
    public required string Purity { get; init; }
    public decimal RatePerGram { get; init; }
    public decimal MakingCharges { get; init; }
    public decimal Discount { get; init; }
    public decimal SubTotal { get; init; }
    public decimal CgstRate { get; init; }
    public decimal SgstRate { get; init; }
    public decimal IgstRate { get; init; }
    public decimal GstAmount { get; init; }
    public decimal TotalAmount { get; init; }
    public required string Status { get; init; }  // PAID | PENDING
}

/// <summary>
/// Gold/silver loan against pledged jewellery.
/// </summary>
public sealed record Loan
{
    public required string Id { get; init; }
    public required string CustomerName { get; init; }
    public required string CustomerPhone { get; init; }
    public required string GovId { get; init; }
    public required string MetalType { get; init; }   // GOLD | SILVER
    public required string ProductDescription { get; init; }
    public decimal Weight { get; init; }
    public required string Purity { get; init; }
    public decimal PrincipalAmount { get; init; }
    public decimal InterestRate { get; init; }
    public required string StartDate { get; init; }
    public required string DueDate { get; init; }
    public decimal TotalRepaid { get; init; }
    public required string Status { get; init; }  // ACTIVE | CLOSED | OVERDUE
}
