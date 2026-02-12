using BillingApp.Models;
using BillingApp.Services;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Printing;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace BillingApp;

public partial class MainWindow : Window
{
    private readonly ExcelWriterService _excelService;
    private readonly string _excelPath;

    private ObservableCollection<Customer> _customers = new();
    private ObservableCollection<Invoice> _invoices = new();
    private ObservableCollection<Loan> _loans = new();

    // Holds the invoice built during Preview, so Save can reuse it
    private Invoice? _previewedInvoice;

    public MainWindow()
    {
        InitializeComponent();

        // Store data in LocalAppData — NOT inside OneDrive (avoids sync locking)
        var localData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BillingApp");
        Directory.CreateDirectory(localData);
        _excelPath = Path.Combine(localData, "billing_data.xlsx");
        _excelService = new ExcelWriterService(_excelPath);

        ExcelPathText.Text = $"📄 {_excelPath}";

        // Excel connectivity check
        if (_excelService.TestConnection(out var connErr))
        {
            SetStatus("✅ Excel connected — Ready", Brushes.LimeGreen);
        }
        else
        {
            SetStatus($"❌ Excel error: {connErr}", Brushes.Red);
        }

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

    /// <summary>Build an Invoice object from the form fields.</summary>
    private Invoice? BuildInvoiceFromForm()
    {
        if (string.IsNullOrWhiteSpace(InvCustId.Text) ||
            string.IsNullOrWhiteSpace(InvItem.Text) ||
            !decimal.TryParse(InvWeight.Text, out var weight) ||
            !decimal.TryParse(InvRate.Text, out var rate))
        {
            SetStatus("❌ Please fill all required fields (Customer ID, Item, Weight, Rate)", Brushes.Red);
            return null;
        }

        decimal.TryParse(InvMaking.Text, out var making);
        decimal.TryParse(InvDiscount.Text, out var discount);

        var subTotal = (weight * rate) + making - discount;
        var cgst = Math.Round(subTotal * 0.015m, 2);
        var sgst = cgst;
        var gstAmount = cgst + sgst;
        var total = subTotal + gstAmount;

        // Return calculation
        decimal returnWeight = 0, returnAmount = 0;
        decimal.TryParse(InvReturnWeight.Text, out returnWeight);
        decimal.TryParse(InvReturnRate.Text, out var returnRate);
        decimal.TryParse(InvReturnAmount.Text, out returnAmount);

        if (ReturnByWeight.IsChecked == true && returnWeight > 0 && returnRate > 0)
        {
            returnAmount = returnWeight * returnRate;
        }

        var netAmount = total - returnAmount;

        return new Invoice
        {
            Id = _excelService.GetNextId("Invoices", "INV-"),
            CustomerId = InvCustId.Text.Trim(),
            CustomerAddress = InvAddress.Text.Trim(),
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
            ReturnWeight = returnWeight,
            ReturnAmount = returnAmount,
            NetAmount = netAmount,
            Status = (InvStatus.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PENDING"
        };
    }

    /// <summary>Preview and print invoice without saving to Excel.</summary>
    private void PreviewInvoice_Click(object sender, RoutedEventArgs e)
    {
        var invoice = BuildInvoiceFromForm();
        if (invoice == null) return;

        _previewedInvoice = invoice;

        var doc = BuildInvoiceDocument(invoice);

        // Show print dialog
        var pd = new PrintDialog();
        if (pd.ShowDialog() == true)
        {
            var paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;
            pd.PrintDocument(paginator, $"Invoice {invoice.Id}");
            SetStatus($"🖨️ Invoice {invoice.Id} printed! Click 'Confirm & Save' to save to Excel.", Brushes.Cyan);
        }
        else
        {
            SetStatus($"🖨️ Invoice {invoice.Id} previewed (not printed). Click 'Confirm & Save' to save to Excel, or Clear to discard.", Brushes.Cyan);
        }

        InvCalcPreview.Text = $"Total: ₹{invoice.TotalAmount:N2}  |  Return: ₹{invoice.ReturnAmount:N2}  |  Net: ₹{invoice.NetAmount:N2}";
    }

    /// <summary>Build a professional Gold Tax Invoice FlowDocument.</summary>
    private static FlowDocument BuildInvoiceDocument(Invoice inv)
    {
        var doc = new FlowDocument
        {
            PageWidth = 680,
            PagePadding = new Thickness(24),
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 10
        };

        var borderBrush = Brushes.Black;
        var headerBg = new SolidColorBrush(Color.FromRgb(240, 230, 200));  // warm cream
        var lightBg = new SolidColorBrush(Color.FromRgb(250, 248, 240));

        // ═══════════════════════════════════════════════════════════════
        // 1. TITLE BAR - "GOLD TAX INVOICE"
        // ═══════════════════════════════════════════════════════════════
        doc.Blocks.Add(new Paragraph(new Run("GOLD TAX INVOICE"))
        {
            FontSize = 14, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 0, 0, 6), Foreground = Brushes.Black
        });

        // ═══════════════════════════════════════════════════════════════
        // 2. SHOP HEADER
        // ═══════════════════════════════════════════════════════════════
        doc.Blocks.Add(new Paragraph(new Run("PHOOL CHANDRA SARAF"))
        {
            FontSize = 22, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Center,
            Foreground = new SolidColorBrush(Color.FromRgb(139, 101, 8)),
            Margin = new Thickness(0, 0, 0, 0)
        });

        doc.Blocks.Add(new Paragraph(new Run("ASHISH JEWELLERS"))
        {
            FontSize = 15, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Center,
            Foreground = Brushes.Black, Margin = new Thickness(0, 0, 0, 2)
        });

        var subP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 0, 0, 2), FontSize = 9 };
        subP.Inlines.Add(new Run("Gold & Silver Ornament Traders") { FontStyle = FontStyles.Italic, Foreground = Brushes.DimGray });
        doc.Blocks.Add(subP);

        doc.Blocks.Add(new Paragraph(new Run("📍 Koraon, Allahabad, Uttar Pradesh - 212306  |  📞 7985494707"))
        {
            FontSize = 9, TextAlignment = TextAlignment.Center,
            Foreground = Brushes.DimGray, Margin = new Thickness(0, 0, 0, 8)
        });

        // ═══════════════════════════════════════════════════════════════
        // 3. INVOICE INFO TABLE (Invoice No, Date, Bill Type)
        // ═══════════════════════════════════════════════════════════════
        var infoTable = new Table { CellSpacing = 0 };
        infoTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        infoTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var infoGroup = new TableRowGroup();

        AddInfoRow(infoGroup, "Invoice No.", inv.Id, "Dated", inv.Date);
        AddInfoRow(infoGroup, "Bill Type", inv.BillType, "Status", inv.Status);

        infoTable.RowGroups.Add(infoGroup);
        SetTableBorder(infoTable);
        doc.Blocks.Add(infoTable);

        // ═══════════════════════════════════════════════════════════════
        // 4. CUSTOMER DETAILS (Billed To)
        // ═══════════════════════════════════════════════════════════════
        var custP = new Paragraph { Margin = new Thickness(0, 6, 0, 6), FontSize = 10 };
        custP.Inlines.Add(new Run("Details of Receiver (Billed To)") { FontWeight = FontWeights.Bold, TextDecorations = TextDecorations.Underline });
        doc.Blocks.Add(custP);

        var custInfo = new Paragraph { Margin = new Thickness(0, 0, 0, 6), FontSize = 10 };
        custInfo.Inlines.Add(new Run($"Customer: ") { Foreground = Brushes.Gray });
        custInfo.Inlines.Add(new Run(inv.CustomerId) { FontWeight = FontWeights.SemiBold });
        if (!string.IsNullOrWhiteSpace(inv.CustomerAddress))
        {
            custInfo.Inlines.Add(new Run($"\nAddress: ") { Foreground = Brushes.Gray });
            custInfo.Inlines.Add(new Run(inv.CustomerAddress));
        }
        doc.Blocks.Add(custInfo);

        // ═══════════════════════════════════════════════════════════════
        // 5. ITEM TABLE
        // ═══════════════════════════════════════════════════════════════
        var itemTable = new Table { CellSpacing = 0 };
        // Columns: Sr | Product & HSN | Purity | Weight(g) | Rate/g | Making | Taxable Amount
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(30) });    // Sr
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) }); // Product
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(55) });    // Purity
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });    // Weight
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });    // Rate
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });    // Making
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(90) });    // Amount

        var itemGroup = new TableRowGroup();

        // Header row
        var headerRow = new TableRow { Background = headerBg };
        AddCell(headerRow, "Sr", FontWeights.Bold, TextAlignment.Center);
        AddCell(headerRow, "Product Name", FontWeights.Bold, TextAlignment.Left);
        AddCell(headerRow, "Purity", FontWeights.Bold, TextAlignment.Center);
        AddCell(headerRow, "Weight(g)", FontWeights.Bold, TextAlignment.Right);
        AddCell(headerRow, "Rate/g", FontWeights.Bold, TextAlignment.Right);
        AddCell(headerRow, "Making", FontWeights.Bold, TextAlignment.Right);
        AddCell(headerRow, "Taxable Amt", FontWeights.Bold, TextAlignment.Right);
        itemGroup.Rows.Add(headerRow);

        // Data row
        var dataRow = new TableRow { Background = lightBg };
        AddCell(dataRow, "1", FontWeights.Normal, TextAlignment.Center);
        AddCell(dataRow, $"{inv.ItemDescription}\n({inv.Metal})", FontWeights.Normal, TextAlignment.Left);
        AddCell(dataRow, inv.Purity, FontWeights.Normal, TextAlignment.Center);
        AddCell(dataRow, $"{inv.Weight:N3}", FontWeights.Normal, TextAlignment.Right);
        AddCell(dataRow, $"{inv.RatePerGram:N2}", FontWeights.Normal, TextAlignment.Right);
        AddCell(dataRow, $"{inv.MakingCharges:N2}", FontWeights.Normal, TextAlignment.Right);
        AddCell(dataRow, $"{inv.SubTotal:N2}", FontWeights.Normal, TextAlignment.Right);
        itemGroup.Rows.Add(dataRow);

        // Totals row
        var totalRow = new TableRow { Background = headerBg };
        AddCell(totalRow, "", FontWeights.Normal, TextAlignment.Center);
        AddCell(totalRow, "Total Pos: 1", FontWeights.Bold, TextAlignment.Left);
        AddCell(totalRow, "", FontWeights.Normal, TextAlignment.Center);
        AddCell(totalRow, $"{inv.Weight:N3}", FontWeights.Bold, TextAlignment.Right);
        AddCell(totalRow, "", FontWeights.Normal, TextAlignment.Right);
        AddCell(totalRow, "", FontWeights.Normal, TextAlignment.Right);
        AddCell(totalRow, $"{inv.SubTotal:N2}", FontWeights.Bold, TextAlignment.Right);
        itemGroup.Rows.Add(totalRow);

        itemTable.RowGroups.Add(itemGroup);
        SetTableBorder(itemTable);
        doc.Blocks.Add(itemTable);

        // ═══════════════════════════════════════════════════════════════
        // 6. GST & GRAND TOTAL TABLE
        // ═══════════════════════════════════════════════════════════════
        var gstTable = new Table { CellSpacing = 0 };
        gstTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        gstTable.Columns.Add(new TableColumn { Width = new GridLength(120) });
        var gstGroup = new TableRowGroup();

        // Discount row (if any)
        if (inv.Discount > 0)
        {
            var discRow = new TableRow();
            AddCell(discRow, "", FontWeights.Normal, TextAlignment.Left);
            AddAlignedValueCell(discRow, $"Discount: -{inv.Discount:N2}");
            gstGroup.Rows.Add(discRow);
        }

        // CGST
        var cgstRow = new TableRow();
        AddCell(cgstRow, "", FontWeights.Normal, TextAlignment.Left);
        AddAlignedValueCell(cgstRow, $"CGST @ {inv.CgstRate}%:  {(inv.GstAmount / 2):N2}");
        gstGroup.Rows.Add(cgstRow);

        // SGST
        var sgstRow = new TableRow();
        AddCell(sgstRow, "", FontWeights.Normal, TextAlignment.Left);
        AddAlignedValueCell(sgstRow, $"SGST @ {inv.SgstRate}%:  {(inv.GstAmount / 2):N2}");
        gstGroup.Rows.Add(sgstRow);

        // Return (if any)
        if (inv.ReturnAmount > 0)
        {
            var retRow = new TableRow();
            var retLabel = inv.ReturnWeight > 0
                ? $"Return ({inv.ReturnWeight:N3}g)"
                : "Return Adjustment";
            AddCell(retRow, retLabel, FontWeights.Bold, TextAlignment.Left);
            AddAlignedValueCell(retRow, $"-{inv.ReturnAmount:N2}");
            gstGroup.Rows.Add(retRow);
        }

        // Grand Total
        var grandRow = new TableRow { Background = headerBg };
        var grandLabel = new TableCell(new Paragraph(new Run("Grand Total :"))
        {
            FontSize = 13, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Margin = new Thickness(4)
        });
        grandRow.Cells.Add(grandLabel);
        var grandValue = new TableCell(new Paragraph(new Run($"₹ {inv.NetAmount:N2}"))
        {
            FontSize = 13, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Margin = new Thickness(4)
        });
        grandRow.Cells.Add(grandValue);
        gstGroup.Rows.Add(grandRow);

        gstTable.RowGroups.Add(gstGroup);
        SetTableBorder(gstTable);
        doc.Blocks.Add(gstTable);

        // ═══════════════════════════════════════════════════════════════
        // 7. AMOUNT IN WORDS
        // ═══════════════════════════════════════════════════════════════
        var wordsP = new Paragraph { Margin = new Thickness(0, 4, 0, 6), FontSize = 9 };
        wordsP.Inlines.Add(new Run("In Words: ") { FontWeight = FontWeights.Bold });
        wordsP.Inlines.Add(new Run(NumberToWords((long)Math.Round(inv.NetAmount)) + " Rupees Only") { FontStyle = FontStyles.Italic });
        doc.Blocks.Add(wordsP);

        // ═══════════════════════════════════════════════════════════════
        // 8. TERMS & CONDITIONS
        // ═══════════════════════════════════════════════════════════════
        doc.Blocks.Add(new Paragraph(new Run("Terms & Conditions:"))
        {
            FontSize = 8, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 4, 0, 2)
        });

        var terms = new List { MarkerStyle = TextMarkerStyle.Decimal, FontSize = 7, Foreground = Brushes.DimGray, Margin = new Thickness(16, 0, 0, 6) };
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Check the goods and weight while buying the jewellery."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Bill should be brought along with the jewellery in case of return or exchange."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Labour and Taxes are non refundable."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Goods once sold will not be exchanged or returned."))));
        doc.Blocks.Add(terms);

        // ═══════════════════════════════════════════════════════════════
        // 9. SIGNATURE AREA
        // ═══════════════════════════════════════════════════════════════
        var sigTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 16, 0, 0) };
        sigTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        sigTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var sigGroup = new TableRowGroup();
        var sigRow = new TableRow();

        var leftSig = new TableCell(new Paragraph(new Run("Receiver's Signature"))
        {
            FontSize = 9, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Left, Margin = new Thickness(0, 20, 0, 0)
        });
        sigRow.Cells.Add(leftSig);

        var rightSig = new TableCell();
        var rightStack = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0) };
        rightStack.Inlines.Add(new Run("For, ASHISH JEWELLERS\n") { FontWeight = FontWeights.Bold, FontSize = 10 });
        rightStack.Inlines.Add(new Run("\n\n"));
        rightStack.Inlines.Add(new Run("(Authorised Signatory)") { FontSize = 8, Foreground = Brushes.Gray });
        rightSig.Blocks.Add(rightStack);
        sigRow.Cells.Add(rightSig);

        sigGroup.Rows.Add(sigRow);
        sigTable.RowGroups.Add(sigGroup);
        doc.Blocks.Add(sigTable);

        return doc;
    }

    // ─── Invoice Document Helpers ──────────────────────────────────────

    private static void AddInfoRow(TableRowGroup group, string lbl1, string val1, string lbl2, string val2)
    {
        var row = new TableRow();
        var cell1 = new TableCell();
        var p1 = new Paragraph { Margin = new Thickness(4, 2, 4, 2) };
        p1.Inlines.Add(new Run($"{lbl1}: ") { Foreground = Brushes.Gray, FontSize = 9 });
        p1.Inlines.Add(new Run(val1) { FontWeight = FontWeights.SemiBold, FontSize = 10 });
        cell1.Blocks.Add(p1);
        cell1.BorderBrush = Brushes.Black;
        cell1.BorderThickness = new Thickness(0.5);
        row.Cells.Add(cell1);

        var cell2 = new TableCell();
        var p2 = new Paragraph { Margin = new Thickness(4, 2, 4, 2), TextAlignment = TextAlignment.Right };
        p2.Inlines.Add(new Run($"{lbl2}: ") { Foreground = Brushes.Gray, FontSize = 9 });
        p2.Inlines.Add(new Run(val2) { FontWeight = FontWeights.SemiBold, FontSize = 10 });
        cell2.Blocks.Add(p2);
        cell2.BorderBrush = Brushes.Black;
        cell2.BorderThickness = new Thickness(0.5);
        row.Cells.Add(cell2);

        group.Rows.Add(row);
    }

    private static void AddCell(TableRow row, string text, FontWeight weight, TextAlignment align)
    {
        var cell = new TableCell(new Paragraph(new Run(text))
        {
            FontWeight = weight, TextAlignment = align, Margin = new Thickness(4, 3, 4, 3), FontSize = 9
        });
        cell.BorderBrush = Brushes.Black;
        cell.BorderThickness = new Thickness(0.5);
        row.Cells.Add(cell);
    }

    private static void AddAlignedValueCell(TableRow row, string text)
    {
        var cell = new TableCell(new Paragraph(new Run(text))
        {
            TextAlignment = TextAlignment.Right, Margin = new Thickness(4, 2, 4, 2), FontSize = 10,
            FontWeight = FontWeights.SemiBold
        });
        cell.BorderBrush = Brushes.Black;
        cell.BorderThickness = new Thickness(0.5);
        row.Cells.Add(cell);
    }

    private static void SetTableBorder(Table table)
    {
        foreach (var rg in table.RowGroups)
            foreach (var r in rg.Rows)
                foreach (var c in r.Cells)
                {
                    c.BorderBrush = Brushes.Black;
                    c.BorderThickness = new Thickness(0.5);
                }
    }

    /// <summary>Convert a number to Indian English words.</summary>
    private static string NumberToWords(long n)
    {
        if (n == 0) return "Zero";
        if (n < 0) return "Minus " + NumberToWords(-n);

        var parts = new System.Collections.Generic.List<string>();

        if (n / 10000000 > 0) { parts.Add(NumberToWords(n / 10000000) + " Crore"); n %= 10000000; }
        if (n / 100000 > 0) { parts.Add(NumberToWords(n / 100000) + " Lakh"); n %= 100000; }
        if (n / 1000 > 0) { parts.Add(NumberToWords(n / 1000) + " Thousand"); n %= 1000; }
        if (n / 100 > 0) { parts.Add(NumberToWords(n / 100) + " Hundred"); n %= 100; }

        if (n > 0)
        {
            var ones = new[] { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
                "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
            var tens = new[] { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

            if (n < 20) parts.Add(ones[n]);
            else
            {
                var s = tens[n / 10];
                if (n % 10 > 0) s += " " + ones[n % 10];
                parts.Add(s);
            }
        }

        return string.Join(" ", parts);
    }

    /// <summary>Save the invoice to Excel. Uses the previewed invoice if available.</summary>
    private async void SaveInvoice_Click(object sender, RoutedEventArgs e)
    {
        var invoice = _previewedInvoice ?? BuildInvoiceFromForm();
        if (invoice == null) return;

        try
        {
            SetStatus("⏳ Calculating GST and writing to Excel...", Brushes.Yellow);

            var count = await _excelService.WriteInvoiceAsync(invoice);
            _invoices.Add(invoice);

            _previewedInvoice = null;
            ClearInvoiceForm();
            UpdateRecordCount();
            SetStatus($"✅ Invoice {invoice.Id} saved! Net: ₹{invoice.NetAmount:N2} (Total: ₹{invoice.TotalAmount:N2}, Return: ₹{invoice.ReturnAmount:N2})", Brushes.LimeGreen);
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
        InvAddress.Text = "";
        InvItem.Text = "";
        InvWeight.Text = "";
        InvRate.Text = "";
        InvMaking.Text = "0";
        InvDiscount.Text = "0";
        InvReturnWeight.Text = "0";
        InvReturnRate.Text = "0";
        InvReturnAmount.Text = "0";
        InvCalcPreview.Text = "";
        _previewedInvoice = null;
        ReturnByWeight.IsChecked = true;
        InvCustId.Focus();
    }

    private void ReturnMode_Changed(object sender, RoutedEventArgs e)
    {
        // Enable/disable fields based on return mode
        if (InvReturnWeight == null || InvReturnRate == null || InvReturnAmount == null) return;

        if (ReturnByWeight.IsChecked == true)
        {
            InvReturnWeight.IsEnabled = true;
            InvReturnRate.IsEnabled = true;
            InvReturnAmount.IsEnabled = false;
        }
        else
        {
            InvReturnWeight.IsEnabled = false;
            InvReturnRate.IsEnabled = false;
            InvReturnAmount.IsEnabled = true;
        }
    }

    // ─── LOAN ──────────────────────────────────────────────────────────

    private async void SaveLoan_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(LoanCustName.Text) ||
            string.IsNullOrWhiteSpace(LoanPhone.Text) ||
            string.IsNullOrWhiteSpace(LoanAddress.Text) ||
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

            var startDate = string.IsNullOrWhiteSpace(LoanStartDate.Text)
                ? DateTime.Now.ToString("yyyy-MM-dd")
                : LoanStartDate.Text.Trim();

            var loan = new Loan
            {
                Id = _excelService.GetNextId("Loans", "L-"),
                CustomerName = LoanCustName.Text.Trim(),
                CustomerPhone = LoanPhone.Text.Trim(),
                CustomerAddress = LoanAddress.Text.Trim(),
                GovId = LoanGovId.Text.Trim(),
                MetalType = (LoanMetal.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "GOLD",
                ProductDescription = LoanProduct.Text.Trim(),
                Weight = weight,
                Purity = (LoanPurity.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "22K",
                PrincipalAmount = principal,
                InterestRate = interest,
                StartDate = startDate,
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
        LoanAddress.Text = "";
        LoanGovId.Text = "";
        LoanProduct.Text = "";
        LoanWeight.Text = "";
        LoanPrincipal.Text = "";
        LoanInterest.Text = "1.5";
        LoanStartDate.Text = "";
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
