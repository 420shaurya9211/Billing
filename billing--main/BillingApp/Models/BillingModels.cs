namespace BillingApp.Models;

/// <summary>
/// Jewellery shop customer.
/// </summary>
public class Customer
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Phone { get; set; } = "";
    public string Address { get; set; } = "";
    public string Gstin { get; set; } = "";
    public decimal TotalPurchases { get; set; }
    public int ActiveLoans { get; set; }
    public int LoyaltyPoints { get; set; }
    public string JoinDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
}

/// <summary>
/// Jewellery invoice with GST breakdown.
/// </summary>
public class Invoice
{
    public string Id { get; set; } = "";
    public string CustomerId { get; set; } = "";
    public string CustomerAddress { get; set; } = "";
    public string Date { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
    public string BillType { get; set; } = "PAKKA";    // PAKKA | KACHA
    public string ItemDescription { get; set; } = "";
    public string Metal { get; set; } = "GOLD";         // GOLD | SILVER
    public decimal Weight { get; set; }
    public string Purity { get; set; } = "22K";
    public decimal RatePerGram { get; set; }
    public decimal MakingCharges { get; set; }
    public decimal Discount { get; set; }
    public decimal SubTotal { get; set; }
    public decimal CgstRate { get; set; } = 1.5m;
    public decimal SgstRate { get; set; } = 1.5m;
    public decimal IgstRate { get; set; }
    public decimal GstAmount { get; set; }
    public decimal TotalAmount { get; set; }
    public decimal ReturnWeight { get; set; }
    public decimal ReturnAmount { get; set; }
    public decimal NetAmount { get; set; }
    public string Status { get; set; } = "PENDING";     // PAID | PENDING
}

/// <summary>
/// Gold/silver loan against pledged jewellery.
/// </summary>
public class Loan
{
    public string Id { get; set; } = "";
    public string CustomerName { get; set; } = "";
    public string CustomerPhone { get; set; } = "";
    public string CustomerAddress { get; set; } = "";
    public string GovId { get; set; } = "";
    public string MetalType { get; set; } = "GOLD";     // GOLD | SILVER
    public string ProductDescription { get; set; } = "";
    public decimal Weight { get; set; }
    public string Purity { get; set; } = "22K";
    public decimal PrincipalAmount { get; set; }
    public decimal InterestRate { get; set; }
    public string StartDate { get; set; } = DateTime.Now.ToString("yyyy-MM-dd");
    public decimal TotalRepaid { get; set; }
    public string Status { get; set; } = "ACTIVE";      // ACTIVE | CLOSED | OVERDUE
}
