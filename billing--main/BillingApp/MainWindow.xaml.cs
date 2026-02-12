using BillingApp.Models;
using BillingApp.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace BillingApp;

public partial class MainWindow : Window
{
    private readonly ExcelWriterService _excelService;
    private readonly string _excelPath;

    private ObservableCollection<Customer> _customers = new();
    private ObservableCollection<Invoice> _invoices = new();
    private ObservableCollection<Loan> _loans = new();

    public MainWindow()
    {
        InitializeComponent();

        _excelPath = Path.Combine(AppContext.BaseDirectory, "billing_data.xlsx");
        _excelService = new ExcelWriterService(_excelPath);

        ExcelPathText.Text = $"📄 {_excelPath}";

        LoadAllData();
    }

    // ─── DATA LOADING ──────────────────────────────────────────────────

    private void LoadAllData()
    {
        try
        {
            _customers = new ObservableCollection<Customer>(_excelService.LoadCustomers());
            _invoices = new ObservableCollection<Invoice>(_excelService.LoadInvoices());
            _loans = new ObservableCollection<Loan>(_excelService.LoadLoans());

            CustomerGrid.ItemsSource = _customers;
            InvoiceGrid.ItemsSource = _invoices;
            LoanGrid.ItemsSource = _loans;

            UpdateRecordCount();
            SetStatus("✅ Data loaded from Excel", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"⚠️ Load error: {ex.Message}", Brushes.Orange);
        }
    }

    // ─── CUSTOMER ──────────────────────────────────────────────────────

    private async void SaveCustomer_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(CustName.Text) ||
            string.IsNullOrWhiteSpace(CustPhone.Text) ||
            string.IsNullOrWhiteSpace(CustAddress.Text))
        {
            SetStatus("❌ Please fill all required fields (Name, Phone, Address)", Brushes.Red);
            return;
        }

        try
        {
            SetStatus("⏳ Writing to Excel...", Brushes.Yellow);

            var customer = new Customer
            {
                Id = _excelService.GetNextId("Customers", "C"),
                Name = CustName.Text.Trim(),
                Phone = CustPhone.Text.Trim(),
                Address = CustAddress.Text.Trim(),
                Gstin = CustGstin.Text.Trim(),
                JoinDate = DateTime.Now.ToString("yyyy-MM-dd")
            };

            var count = await _excelService.WriteCustomerAsync(customer);
            _customers.Add(customer);

            ClearCustomerForm();
            UpdateRecordCount();
            SetStatus($"✅ Customer '{customer.Name}' saved to Excel! (Row #{count})", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Error: {ex.Message}", Brushes.Red);
        }
    }

    private void ClearCustomer_Click(object sender, RoutedEventArgs e) => ClearCustomerForm();

    private void ClearCustomerForm()
    {
        CustName.Text = "";
        CustPhone.Text = "";
        CustAddress.Text = "";
        CustGstin.Text = "";
        CustName.Focus();
    }

    // ─── INVOICE ───────────────────────────────────────────────────────

    private async void SaveInvoice_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(InvCustId.Text) ||
            string.IsNullOrWhiteSpace(InvItem.Text) ||
            !decimal.TryParse(InvWeight.Text, out var weight) ||
            !decimal.TryParse(InvRate.Text, out var rate))
        {
            SetStatus("❌ Please fill all required fields (Customer ID, Item, Weight, Rate)", Brushes.Red);
            return;
        }

        try
        {
            SetStatus("⏳ Calculating GST and writing to Excel...", Brushes.Yellow);

            decimal.TryParse(InvMaking.Text, out var making);
            decimal.TryParse(InvDiscount.Text, out var discount);

            var subTotal = (weight * rate) + making - discount;
            var cgst = Math.Round(subTotal * 0.015m, 2);
            var sgst = cgst;
            var gstAmount = cgst + sgst;
            var total = subTotal + gstAmount;

            var invoice = new Invoice
            {
                Id = _excelService.GetNextId("Invoices", "INV-"),
                CustomerId = InvCustId.Text.Trim(),
                Date = DateTime.Now.ToString("yyyy-MM-dd"),
                BillType = (InvBillType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PAKKA",
                ItemDescription = InvItem.Text.Trim(),
                Metal = (InvMetal.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "GOLD",
                Weight = weight,
                Purity = (InvPurity.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "22K",
                RatePerGram = rate,
                MakingCharges = making,
                Discount = discount,
                SubTotal = subTotal,
                CgstRate = 1.5m,
                SgstRate = 1.5m,
                IgstRate = 0,
                GstAmount = gstAmount,
                TotalAmount = total,
                Status = (InvStatus.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PENDING"
            };

            var count = await _excelService.WriteInvoiceAsync(invoice);
            _invoices.Add(invoice);

            ClearInvoiceForm();
            UpdateRecordCount();
            SetStatus($"✅ Invoice {invoice.Id} saved! Total: ₹{total:N2} (incl. GST ₹{gstAmount:N2})", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Error: {ex.Message}", Brushes.Red);
        }
    }

    private void ClearInvoice_Click(object sender, RoutedEventArgs e) => ClearInvoiceForm();

    private void ClearInvoiceForm()
    {
        InvCustId.Text = "";
        InvItem.Text = "";
        InvWeight.Text = "";
        InvRate.Text = "";
        InvMaking.Text = "0";
        InvDiscount.Text = "0";
        InvCalcPreview.Text = "";
        InvCustId.Focus();
    }

    // ─── LOAN ──────────────────────────────────────────────────────────

    private async void SaveLoan_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(LoanCustName.Text) ||
            string.IsNullOrWhiteSpace(LoanPhone.Text) ||
            string.IsNullOrWhiteSpace(LoanGovId.Text) ||
            string.IsNullOrWhiteSpace(LoanProduct.Text) ||
            !decimal.TryParse(LoanWeight.Text, out var weight) ||
            !decimal.TryParse(LoanPrincipal.Text, out var principal))
        {
            SetStatus("❌ Please fill all required fields", Brushes.Red);
            return;
        }

        try
        {
            SetStatus("⏳ Writing loan to Excel...", Brushes.Yellow);

            decimal.TryParse(LoanInterest.Text, out var interest);
            var dueDate = string.IsNullOrWhiteSpace(LoanDueDate.Text)
                ? DateTime.Now.AddMonths(6).ToString("yyyy-MM-dd")
                : LoanDueDate.Text.Trim();

            var loan = new Loan
            {
                Id = _excelService.GetNextId("Loans", "L-"),
                CustomerName = LoanCustName.Text.Trim(),
                CustomerPhone = LoanPhone.Text.Trim(),
                GovId = LoanGovId.Text.Trim(),
                MetalType = (LoanMetal.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "GOLD",
                ProductDescription = LoanProduct.Text.Trim(),
                Weight = weight,
                Purity = (LoanPurity.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "22K",
                PrincipalAmount = principal,
                InterestRate = interest,
                StartDate = DateTime.Now.ToString("yyyy-MM-dd"),
                DueDate = dueDate,
                TotalRepaid = 0,
                Status = "ACTIVE"
            };

            var count = await _excelService.WriteLoanAsync(loan);
            _loans.Add(loan);

            ClearLoanForm();
            UpdateRecordCount();
            SetStatus($"✅ Loan {loan.Id} for {loan.CustomerName} saved! Principal: ₹{principal:N2}", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Error: {ex.Message}", Brushes.Red);
        }
    }

    private void ClearLoan_Click(object sender, RoutedEventArgs e) => ClearLoanForm();

    private void ClearLoanForm()
    {
        LoanCustName.Text = "";
        LoanPhone.Text = "";
        LoanGovId.Text = "";
        LoanProduct.Text = "";
        LoanWeight.Text = "";
        LoanPrincipal.Text = "";
        LoanInterest.Text = "1.5";
        LoanDueDate.Text = "";
        LoanCustName.Focus();
    }

    // ─── TOOLBAR ───────────────────────────────────────────────────────

    private void OpenExcel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (File.Exists(_excelPath))
            {
                Process.Start(new ProcessStartInfo(_excelPath) { UseShellExecute = true });
                SetStatus("📂 Opened Excel file", Brushes.Cyan);
            }
            else
            {
                SetStatus("⚠️ Excel file not found — save some data first!", Brushes.Orange);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"❌ Cannot open Excel: {ex.Message}", Brushes.Red);
        }
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => LoadAllData();

    private void Tab_Changed(object sender, SelectionChangedEventArgs e) { }

    // ─── HELPERS ───────────────────────────────────────────────────────

    private void SetStatus(string message, Brush color)
    {
        StatusText.Text = message;
        StatusText.Foreground = color;
    }

    private void UpdateRecordCount()
    {
        RecordCountText.Text = $"Customers: {_customers.Count}  |  Invoices: {_invoices.Count}  |  Loans: {_loans.Count}";
    }

    protected override void OnClosed(EventArgs e)
    {
        _excelService.Dispose();
        base.OnClosed(e);
    }
}
