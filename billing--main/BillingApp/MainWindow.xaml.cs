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

    // Multi-item collections for the add-item sub-forms
    private readonly ObservableCollection<InvoiceItem> _invoiceItems = new();
    private readonly ObservableCollection<LoanItem> _loanItems = new();

    // Holds the invoice built during Preview, so Save can reuse it
    private Invoice? _previewedInvoice;

    public MainWindow()
    {
        InitializeComponent();

        // Bind item collections to grids
        InvItemsGrid.ItemsSource = _invoiceItems;
        LoanItemsGrid.ItemsSource = _loanItems;

        // Store data in LocalAppData â€” NOT inside OneDrive (avoids sync locking)
        var localData = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BillingApp");
        Directory.CreateDirectory(localData);
        _excelPath = Path.Combine(localData, "billing_data.xlsx");
        _excelService = new ExcelWriterService(_excelPath);

        ExcelPathText.Text = $"ðŸ“„ {_excelPath}";

        // Excel connectivity check
        if (_excelService.TestConnection(out var connErr))
        {
            SetStatus("âœ… Excel connected â€” Ready", Brushes.LimeGreen);
        }
        else
        {
            SetStatus($"âŒ Excel error: {connErr}", Brushes.Red);
        }

        LoadAllData();
    }

    // â”€â”€â”€ DATA LOADING â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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
            SetStatus("âœ… Data loaded from Excel", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"âš ï¸ Load error: {ex.Message}", Brushes.Orange);
        }
    }

    // â”€â”€â”€ CUSTOMER â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    private async void SaveCustomer_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(CustName.Text) ||
            string.IsNullOrWhiteSpace(CustPhone.Text) ||
            string.IsNullOrWhiteSpace(CustAddress.Text))
        {
            SetStatus("âŒ Please fill all required fields (Name, Phone, Address)", Brushes.Red);
            return;
        }

        try
        {
            SetStatus("â³ Writing to Excel...", Brushes.Yellow);

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
            SetStatus($"âœ… Customer '{customer.Name}' saved to Excel! (Row #{count})", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"âŒ Error: {ex.Message}", Brushes.Red);
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

    // â”€â”€â”€ INVOICE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    // â”€â”€ Add/Remove invoice item handlers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    private void AddInvoiceItem_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(InvItem.Text) ||
            !decimal.TryParse(InvWeight.Text, out var w) ||
            !decimal.TryParse(InvRate.Text, out var r))
        {
            SetStatus("âŒ Fill Item, Weight, and Rate to add an item", Brushes.Red);
            return;
        }

        decimal.TryParse(InvMaking.Text, out var m);

        _invoiceItems.Add(new InvoiceItem
        {
            ItemDescription = InvItem.Text.Trim(),
            Metal = (InvMetal.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "GOLD",
            Purity = InvPurity.Text.Trim(),
            Weight = w,
            RatePerGram = r,
            MakingCharges = m
        });

        // Clear sub-form for next item
        InvItem.Text = "";
        InvWeight.Text = "";
        InvRate.Text = "";
        InvMaking.Text = "0";
        InvItem.Focus();
        SetStatus($"âœ… Item added ({_invoiceItems.Count} items total)", Brushes.LimeGreen);
    }

    private void RemoveInvoiceItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is InvoiceItem item)
        {
            _invoiceItems.Remove(item);
            SetStatus($"ðŸ—‘ï¸ Item removed ({_invoiceItems.Count} items remaining)", Brushes.Orange);
        }
    }

    /// <summary>Build an Invoice object from the form fields.</summary>
    private Invoice? BuildInvoiceFromForm()
    {
        if (string.IsNullOrWhiteSpace(InvCustId.Text) || _invoiceItems.Count == 0)
        {
            SetStatus("âŒ Please fill Customer ID and add at least one item", Brushes.Red);
            return null;
        }

        decimal.TryParse(InvDiscount.Text, out var discount);

        var billType = (InvBillType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PAKKA";

        // Sum all items
        var items = _invoiceItems.ToList();
        var totalWeight = items.Sum(i => i.Weight);
        var subTotal = items.Sum(i => i.Amount) - discount;

        // KACHA bills have no GST
        decimal gstAmount = 0, total;
        if (billType == "PAKKA")
        {
            var cgst = Math.Round(subTotal * 0.015m, 2);
            gstAmount = cgst * 2;
            total = subTotal + gstAmount;
        }
        else
        {
            total = subTotal;
        }

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

        // For backward-compat: first item populates legacy single-item fields
        var first = items[0];

        return new Invoice
        {
            Id = _excelService.GetNextId("Invoices", "INV-"),
            CustomerId = InvCustId.Text.Trim(),
            CustomerPhone = InvPhone.Text.Trim(),
            CustomerAddress = InvAddress.Text.Trim(),
            Date = InvDate.SelectedDate?.ToString("yyyy-MM-dd") ?? DateTime.Now.ToString("yyyy-MM-dd"),
            BillType = billType,
            ItemDescription = string.Join(" | ", items.Select(i => i.ItemDescription)),
            Metal = string.Join(" | ", items.Select(i => i.Metal)),
            Weight = totalWeight,
            Purity = string.Join(" | ", items.Select(i => i.Purity)),
            RatePerGram = first.RatePerGram,
            MakingCharges = items.Sum(i => i.MakingCharges),
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
            Status = (InvStatus.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PENDING",
            Items = items
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
            SetStatus($"ðŸ–¨ï¸ Invoice {invoice.Id} printed! Click 'Confirm & Save' to save to Excel.", Brushes.Cyan);
        }
        else
        {
            SetStatus($"ðŸ–¨ï¸ Invoice {invoice.Id} previewed (not printed). Click 'Confirm & Save' to save to Excel, or Clear to discard.", Brushes.Cyan);
        }

        InvCalcPreview.Text = $"Total: â‚¹{invoice.TotalAmount:N2}  |  Return: â‚¹{invoice.ReturnAmount:N2}  |  Net: â‚¹{invoice.NetAmount:N2}";
    }

    /// <summary>Build a professional Gold Tax Invoice FlowDocument (Half-A4 with Hindi).</summary>
    private static FlowDocument BuildInvoiceDocument(Invoice inv)
    {
        var doc = new FlowDocument
        {
            // Half-A4: 210mm Ã— 148.5mm = 794 Ã— 561 px at 96 DPI
            PageWidth = 794,
            PageHeight = 561,
            PagePadding = new Thickness(24, 16, 24, 12),
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 8,
            ColumnWidth = 999,
            IsColumnWidthFlexible = false
        };

        var hindiFont = new FontFamily("Nirmala UI");
        var borderBrush = Brushes.Black;
        var goldColor = new SolidColorBrush(Color.FromRgb(139, 101, 8));
        var headerBg = new SolidColorBrush(Color.FromRgb(245, 235, 210));
        var lightBg = new SolidColorBrush(Color.FromRgb(255, 252, 245));
        var accentBg = new SolidColorBrush(Color.FromRgb(250, 245, 225));
        var darkHeaderBg = new SolidColorBrush(Color.FromRgb(85, 65, 20));

        // â•â•â• 1. TITLE BAR â•â•â•
        var titleTable = new Table { CellSpacing = 0 };
        titleTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var titleGroup = new TableRowGroup();
        var titleRow = new TableRow { Background = darkHeaderBg };
        var titleP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        titleP.Inlines.Add(new Run("GOLD TAX INVOICE") { FontSize = 11, FontWeight = FontWeights.Bold, Foreground = Brushes.White });
        titleP.Inlines.Add(new Run("  |  बिल / चालान") { FontSize = 10, FontWeight = FontWeights.Bold, Foreground = Brushes.Gold, FontFamily = hindiFont });
        var titleCell = new TableCell(titleP);
        titleCell.BorderBrush = borderBrush;
        titleCell.BorderThickness = new Thickness(1);
        titleCell.Padding = new Thickness(0, 2, 0, 2);
        titleRow.Cells.Add(titleCell);
        titleGroup.Rows.Add(titleRow);
        titleTable.RowGroups.Add(titleGroup);
        doc.Blocks.Add(titleTable);

        // â•â•â• 2. SHOP HEADER (compact) â•â•â•
        var shopTable = new Table { CellSpacing = 0 };
        shopTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var shopGroup = new TableRowGroup();
        var shopRow = new TableRow { Background = accentBg };
        var shopCell = new TableCell();
        shopCell.BorderBrush = borderBrush;
        shopCell.BorderThickness = new Thickness(1, 0, 1, 0);
        shopCell.Padding = new Thickness(4, 4, 4, 4);

        var shopP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0) };
        shopP.Inlines.Add(new Run("PHOOLCHANDRA SARAF JEWELLERS") { FontSize = 16, FontWeight = FontWeights.Bold, Foreground = goldColor });
        shopP.Inlines.Add(new Run("  (फूलचन्द्र सर्राफ ज्वैलर्स)") { FontSize = 11, FontWeight = FontWeights.Bold, Foreground = goldColor, FontFamily = hindiFont });
        shopCell.Blocks.Add(shopP);


        var addrP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 1, 0, 0) };
        addrP.Inlines.Add(new Run("Bahadur Shah Nagar, Koraon-Prayagraj  |  Ph: 7518318070") { FontSize = 7, Foreground = Brushes.DimGray });
        shopCell.Blocks.Add(addrP);

        shopRow.Cells.Add(shopCell);
        shopGroup.Rows.Add(shopRow);
        shopTable.RowGroups.Add(shopGroup);
        doc.Blocks.Add(shopTable);

        // â•â•â• 3. META + CUSTOMER (side by side) â•â•â•
        var infoTable = new Table { CellSpacing = 0 };
        infoTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        infoTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var infoGroup = new TableRowGroup();
        var infoRow = new TableRow();

        // Left: Invoice details
        var invCell = new TableCell();
        invCell.BorderBrush = borderBrush;
        invCell.BorderThickness = new Thickness(1, 0, 0.5, 1);
        invCell.Padding = new Thickness(6, 3, 6, 3);
        var invInfo = new Paragraph { FontSize = 8, LineHeight = 13, Margin = new Thickness(0) };
        invInfo.Inlines.Add(new Run("Invoice No: ") { Foreground = Brushes.Gray });
        invInfo.Inlines.Add(new Run(inv.Id) { FontWeight = FontWeights.Bold });
        invInfo.Inlines.Add(new Run($"\nDate: {inv.Date}"));
        invInfo.Inlines.Add(new Run($"\nBill Type: {inv.BillType}") { FontWeight = FontWeights.SemiBold });
        invInfo.Inlines.Add(new Run($"\nStatus: {inv.Status}"));
        invCell.Blocks.Add(invInfo);
        infoRow.Cells.Add(invCell);

        // Right: Customer details
        var custCell = new TableCell();
        custCell.BorderBrush = borderBrush;
        custCell.BorderThickness = new Thickness(0.5, 0, 1, 1);
        custCell.Padding = new Thickness(6, 3, 6, 3);
        var custInfo = new Paragraph { FontSize = 8, LineHeight = 13, Margin = new Thickness(0) };
        custInfo.Inlines.Add(new Run("Customer / ग्राहक:") { Foreground = Brushes.Gray, FontFamily = hindiFont });
        custInfo.Inlines.Add(new Run($"\n{inv.CustomerId}") { FontWeight = FontWeights.Bold });
        if (!string.IsNullOrWhiteSpace(inv.CustomerAddress))
            custInfo.Inlines.Add(new Run($"\nAddress: {inv.CustomerAddress}"));
        if (!string.IsNullOrWhiteSpace(inv.CustomerPhone))
            custInfo.Inlines.Add(new Run($"\nMobile: {inv.CustomerPhone}") { FontWeight = FontWeights.SemiBold });
        custCell.Blocks.Add(custInfo);
        infoRow.Cells.Add(custCell);

        infoGroup.Rows.Add(infoRow);
        infoTable.RowGroups.Add(infoGroup);
        doc.Blocks.Add(infoTable);

        // â•â•â• 4. ITEM + GST + GRAND TOTAL TABLE â•â•â•
        var itemTable = new Table { CellSpacing = 0 };
        // 8 columns â€“ rebalanced for half-A4 (746px = 794 - 48 padding)
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(25) });       // Sr
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(175) });      // Product+HSN
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(45) });       // Purity
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(60) });       // Gross Wt
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(60) });       // Net Wt
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });       // Rate
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });       // Making
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(141) });      // Amount

        var itemGroup = new TableRowGroup();

        // Header row
        var hdr = new TableRow { Background = headerBg };
        AddItemCell(hdr, "Sr", FontWeights.Bold, TextAlignment.Center, 7);
        AddItemCell(hdr, "Product / विवरण", FontWeights.Bold, TextAlignment.Left, 7);
        AddItemCell(hdr, "Purity", FontWeights.Bold, TextAlignment.Center, 7);
        AddItemCell(hdr, "Gross Wt", FontWeights.Bold, TextAlignment.Right, 7);
        AddItemCell(hdr, "Net Wt", FontWeights.Bold, TextAlignment.Right, 7);
        AddItemCell(hdr, "Rate/g", FontWeights.Bold, TextAlignment.Right, 7);
        AddItemCell(hdr, "Making", FontWeights.Bold, TextAlignment.Right, 7);
        var amountLabel = inv.BillType == "PAKKA" ? "Taxable Amt" : "Total Amt";
        AddItemCell(hdr, amountLabel, FontWeights.Bold, TextAlignment.Right, 7);
        itemGroup.Rows.Add(hdr);

        // Data rows
        var itemsList = inv.Items != null && inv.Items.Count > 0
            ? inv.Items
            : new List<InvoiceItem> { new InvoiceItem {
                ItemDescription = inv.ItemDescription, Metal = inv.Metal,
                Purity = inv.Purity, Weight = inv.Weight,
                RatePerGram = inv.RatePerGram, MakingCharges = inv.MakingCharges } };

        for (int i = 0; i < itemsList.Count; i++)
        {
            var item = itemsList[i];
            var bg = (i % 2 == 0) ? lightBg : accentBg;
            var dRow = new TableRow { Background = bg };
            AddItemCell(dRow, $"{i + 1}", FontWeights.Normal, TextAlignment.Center, 7);
            AddItemCell(dRow, $"{item.ItemDescription} ({item.Metal})", FontWeights.Normal, TextAlignment.Left, 7);
            AddItemCell(dRow, item.Purity, FontWeights.Normal, TextAlignment.Center, 7);
            AddItemCell(dRow, $"{item.Weight:N3}", FontWeights.Normal, TextAlignment.Right, 7);
            AddItemCell(dRow, $"{item.Weight:N3}", FontWeights.Normal, TextAlignment.Right, 7);
            AddItemCell(dRow, $"{item.RatePerGram:N0}", FontWeights.Normal, TextAlignment.Right, 7);
            AddItemCell(dRow, $"{item.MakingCharges:N1}", FontWeights.Normal, TextAlignment.Right, 7);
            AddItemCell(dRow, $"{item.Amount:N2}", FontWeights.Normal, TextAlignment.Right, 7);
            itemGroup.Rows.Add(dRow);
        }

        // Total Pcs row
        var tRow = new TableRow { Background = headerBg };
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Center, 7);
        AddItemCell(tRow, $"Total Pcs: {itemsList.Count}", FontWeights.Bold, TextAlignment.Left, 7);
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Center, 7);
        AddItemCell(tRow, $"{inv.Weight:N3}", FontWeights.Bold, TextAlignment.Right, 7);
        AddItemCell(tRow, $"{inv.Weight:N3}", FontWeights.Bold, TextAlignment.Right, 7);
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Right, 7);
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Right, 7);
        AddItemCell(tRow, $"{itemsList.Sum(x => x.Amount):N2}", FontWeights.Bold, TextAlignment.Right, 7);
        itemGroup.Rows.Add(tRow);

        // Discount row
        if (inv.Discount > 0)
        {
            var discRow = new TableRow();
            discRow.Cells.Add(MakeSpanCell("Less Discount:", 7, TextAlignment.Right, FontWeights.Normal, 7));
            AddItemCell(discRow, $"-{inv.Discount:N2}", FontWeights.Bold, TextAlignment.Right, 7);
            itemGroup.Rows.Add(discRow);
        }

        // CGST / SGST â€” PAKKA only
        if (inv.BillType == "PAKKA")
        {
            var cgstRow = new TableRow();
            cgstRow.Cells.Add(MakeSpanCell($"CGST @ {inv.CgstRate}%", 7, TextAlignment.Right, FontWeights.Normal, 7));
            AddItemCell(cgstRow, $"{(inv.GstAmount / 2):N2}", FontWeights.Bold, TextAlignment.Right, 7);
            itemGroup.Rows.Add(cgstRow);

            var sgstRow = new TableRow();
            sgstRow.Cells.Add(MakeSpanCell($"SGST @ {inv.SgstRate}%", 7, TextAlignment.Right, FontWeights.Normal, 7));
            AddItemCell(sgstRow, $"{(inv.GstAmount / 2):N2}", FontWeights.Bold, TextAlignment.Right, 7);
            itemGroup.Rows.Add(sgstRow);
        }

        // Return row
        if (inv.ReturnAmount > 0)
        {
            var retLabel = inv.ReturnWeight > 0 ? $"Return ({inv.ReturnWeight:N3}g)" : "Return Adjustment";
            var retRow = new TableRow();
            retRow.Cells.Add(MakeSpanCell(retLabel, 7, TextAlignment.Right, FontWeights.Normal, 7));
            AddItemCell(retRow, $"-{inv.ReturnAmount:N2}", FontWeights.Bold, TextAlignment.Right, 7);
            itemGroup.Rows.Add(retRow);
        }

        // Grand Total row
        var grandRow = new TableRow { Background = darkHeaderBg };
        var wordsText = "In Words: " + NumberToWords((long)Math.Round(inv.NetAmount)) + " Rupees Only";
        var wordsCellGrand = new TableCell(new Paragraph(new Run(wordsText))
        {
            FontSize = 6.5, FontWeight = FontWeights.SemiBold, FontStyle = FontStyles.Italic,
            Foreground = Brushes.White, Margin = new Thickness(2, 2, 2, 2)
        }) { ColumnSpan = 6 };
        wordsCellGrand.BorderBrush = borderBrush;
        wordsCellGrand.BorderThickness = new Thickness(0.5);
        grandRow.Cells.Add(wordsCellGrand);

        var gtLabelCell = new TableCell(new Paragraph(new Run("Grand Total :"))
        {
            FontSize = 9, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(2, 2, 2, 2)
        });
        gtLabelCell.BorderBrush = borderBrush;
        gtLabelCell.BorderThickness = new Thickness(0.5);
        grandRow.Cells.Add(gtLabelCell);

        var gtValCell = new TableCell(new Paragraph(new Run($"â‚¹ {inv.NetAmount:N2}"))
        {
            FontSize = 10, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(2, 2, 2, 2)
        });
        gtValCell.BorderBrush = borderBrush;
        gtValCell.BorderThickness = new Thickness(0.5);
        grandRow.Cells.Add(gtValCell);
        itemGroup.Rows.Add(grandRow);

        itemTable.RowGroups.Add(itemGroup);
        SetTableBorder(itemTable);
        doc.Blocks.Add(itemTable);

        // â•â•â• 5. T&C + BANK + SIGNATURE (compact) â•â•â•
        var footerTable = new Table { CellSpacing = 0 };
        footerTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        footerTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var footerGroup = new TableRowGroup();
        var footerRow = new TableRow();

        // Left: Terms
        var termsCell = new TableCell();
        termsCell.BorderBrush = borderBrush;
        termsCell.BorderThickness = new Thickness(1, 0, 0.5, 1);
        termsCell.Padding = new Thickness(4, 2, 4, 2);

        var termsHeader = new Paragraph { Margin = new Thickness(0, 0, 0, 1) };
        termsHeader.Inlines.Add(new Run("Terms & Conditions") { FontSize = 7, FontWeight = FontWeights.Bold });
        termsHeader.Inlines.Add(new Run(" / à¤¨à¤¿à¤¯à¤® à¤”à¤° à¤¶à¤°à¥à¤¤à¥‡à¤‚") { FontSize = 7, FontWeight = FontWeights.Bold, FontFamily = hindiFont });
        termsCell.Blocks.Add(termsHeader);

        var terms = new List { MarkerStyle = TextMarkerStyle.Decimal, FontSize = 6, Foreground = Brushes.DimGray, Margin = new Thickness(10, 0, 0, 0) };
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Check goods and weight while buying."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Bill must be brought for return/exchange."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Labour & Taxes non-refundable."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("KDM: 15% depreciation | 916 HM: 10%."))));
        termsCell.Blocks.Add(terms);
        footerRow.Cells.Add(termsCell);

        // Right: Bank + Signature
        var rightCell = new TableCell();
        rightCell.BorderBrush = borderBrush;
        rightCell.BorderThickness = new Thickness(0.5, 0, 1, 1);
        rightCell.Padding = new Thickness(4, 2, 4, 2);

        rightCell.Blocks.Add(new Paragraph(new Run("Bank Details:")) { FontSize = 7, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 1) });
        var bankInfo = new Paragraph { FontSize = 6.5, Foreground = Brushes.DimGray, LineHeight = 11 };
        bankInfo.Inlines.Add(new Run("Bank: ________  |  A/c: ________  |  IFSC: ________"));
        rightCell.Blocks.Add(bankInfo);

        var forP = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0, 6, 0, 2) };
        forP.Inlines.Add(new Run("For, ") { FontSize = 7, Foreground = Brushes.Gray });
        forP.Inlines.Add(new Run("PHOOLCHANDRA SARAF JEWELLERS") { FontSize = 8, FontWeight = FontWeights.Bold });
        rightCell.Blocks.Add(forP);
        rightCell.Blocks.Add(new Paragraph(new Run("(Authorised Signatory)")) { FontSize = 6, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Right });
        footerRow.Cells.Add(rightCell);

        footerGroup.Rows.Add(footerRow);
        footerTable.RowGroups.Add(footerGroup);
        doc.Blocks.Add(footerTable);

        // â•â•â• 6. FOOTER NOTE â•â•â•
        doc.Blocks.Add(new Paragraph(new Run("This is a computer generated invoice."))
        {
            FontSize = 6, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 3, 0, 0)
        });

        return doc;
    }


    // â”€â”€â”€ Invoice Document Helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    /// <summary>Add a label-value cell to a meta row.</summary>
    private static void AddLabelValueCell(TableRow row, string label, string value)
    {
        var cell = new TableCell();
        cell.BorderBrush = Brushes.Black;
        cell.BorderThickness = new Thickness(1);
        cell.Padding = new Thickness(6, 3, 6, 3);
        var p = new Paragraph { Margin = new Thickness(0) };
        p.Inlines.Add(new Run($"{label}\n") { FontSize = 8, Foreground = Brushes.Gray });
        p.Inlines.Add(new Run(value) { FontSize = 10, FontWeight = FontWeights.Bold });
        cell.Blocks.Add(p);
        row.Cells.Add(cell);
    }

    /// <summary>Add a cell to the item table with specific font size.</summary>
    private static void AddItemCell(TableRow row, string text, FontWeight weight, TextAlignment align, double fontSize)
    {
        var cell = new TableCell(new Paragraph(new Run(text))
        {
            FontWeight = weight, TextAlignment = align, Margin = new Thickness(4, 3, 4, 3), FontSize = fontSize
        });
        cell.BorderBrush = Brushes.Black;
        cell.BorderThickness = new Thickness(0.5);
        row.Cells.Add(cell);
    }

    /// <summary>Create a TableCell that spans multiple columns (for CGST/SGST/Grand Total labels).</summary>
    private static TableCell MakeSpanCell(string text, int colSpan, TextAlignment align, FontWeight weight, double fontSize)
    {
        var cell = new TableCell(new Paragraph(new Run(text))
        {
            FontWeight = weight, TextAlignment = align, Margin = new Thickness(4, 3, 4, 3), FontSize = fontSize
        })
        {
            ColumnSpan = colSpan
        };
        cell.BorderBrush = Brushes.Black;
        cell.BorderThickness = new Thickness(0.5);
        return cell;
    }

    /// <summary>Add a summary value row (label + value in one cell, right-aligned).</summary>
    private static void AddSummaryValueRow(TableRow row, string label, string value)
    {
        var cell = new TableCell();
        cell.BorderBrush = Brushes.Black;
        cell.BorderThickness = new Thickness(1);
        cell.Padding = new Thickness(6, 2, 6, 2);
        var p = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0) };
        p.Inlines.Add(new Run($"{label}  ") { FontSize = 9, Foreground = Brushes.DimGray });
        p.Inlines.Add(new Run(value) { FontSize = 10, FontWeight = FontWeights.Bold });
        cell.Blocks.Add(p);
        row.Cells.Add(cell);
    }

    /// <summary>Legacy helper kept for compatibility.</summary>
    private static void AddCell(TableRow row, string text, FontWeight weight, TextAlignment align)
    {
        AddItemCell(row, text, weight, align, 9);
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
            SetStatus("â³ Calculating GST and writing to Excel...", Brushes.Yellow);

            var count = await _excelService.WriteInvoiceAsync(invoice);
            _invoices.Add(invoice);

            _previewedInvoice = null;
            ClearInvoiceForm();
            UpdateRecordCount();
            SetStatus($"âœ… Invoice {invoice.Id} saved! Net: â‚¹{invoice.NetAmount:N2} (Total: â‚¹{invoice.TotalAmount:N2}, Return: â‚¹{invoice.ReturnAmount:N2})", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"âŒ Error: {ex.Message}", Brushes.Red);
        }
    }

    private void ClearInvoice_Click(object sender, RoutedEventArgs e) => ClearInvoiceForm();

    private void ClearInvoiceForm()
    {
        InvCustId.Text = "";
        InvPhone.Text = "";
        InvAddress.Text = "";
        InvDate.SelectedDate = DateTime.Today;
        InvItem.Text = "";
        InvWeight.Text = "";
        InvRate.Text = "";
        InvMaking.Text = "0";
        InvDiscount.Text = "0";
        InvPurity.Text = "91.6";
        InvReturnWeight.Text = "0";
        InvReturnRate.Text = "0";
        InvReturnAmount.Text = "0";
        InvCalcPreview.Text = "";
        _previewedInvoice = null;
        _invoiceItems.Clear();
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

    // â”€â”€â”€ LOAN â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    private async void SaveLoan_Click(object sender, RoutedEventArgs e)
    {
        var loan = BuildLoanFromForm();
        if (loan == null) return;

        try
        {
            SetStatus("â³ Writing loan to Excel...", Brushes.Yellow);

            var count = await _excelService.WriteLoanAsync(loan);
            _loans.Add(loan);

            ClearLoanForm();
            UpdateRecordCount();
            SetStatus($"âœ… Loan {loan.Id} for {loan.CustomerName} saved! Principal: â‚¹{loan.PrincipalAmount:N2}", Brushes.LimeGreen);
        }
        catch (Exception ex)
        {
            SetStatus($"âŒ Error: {ex.Message}", Brushes.Red);
        }
    }

    private void ClearLoan_Click(object sender, RoutedEventArgs e) => ClearLoanForm();

    private void ClearLoanForm()
    {
        LoanCustName.Text = "";
        LoanPhone.Text = "";
        LoanAddress.Text = "";
        LoanGovId.Text = "";
        LoanIdType.SelectedIndex = 0;
        LoanProduct.Text = "";
        LoanWeight.Text = "";
        LoanPurity.Text = "91.6";
        LoanPrincipal.Text = "";
        LoanInterest.Text = "1.5";
        LoanStartDatePicker.SelectedDate = DateTime.Today;
        _loanItems.Clear();
        LoanCustName.Focus();
    }

    // â”€â”€â”€ LOAN INVOICE & DUMMY INVOICE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    // â”€â”€ Add/Remove loan item handlers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

    private void AddLoanItem_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(LoanProduct.Text) ||
            !decimal.TryParse(LoanWeight.Text, out var w))
        {
            SetStatus("âŒ Fill Product and Weight to add an item", Brushes.Red);
            return;
        }

        _loanItems.Add(new LoanItem
        {
            ProductDescription = LoanProduct.Text.Trim(),
            MetalType = (LoanMetal.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "GOLD",
            Purity = LoanPurity.Text.Trim(),
            Weight = w
        });

        LoanProduct.Text = "";
        LoanWeight.Text = "";
        LoanProduct.Focus();
        SetStatus($"âœ… Item added ({_loanItems.Count} items total)", Brushes.LimeGreen);
    }

    private void RemoveLoanItem_Click(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement fe && fe.DataContext is LoanItem item)
        {
            _loanItems.Remove(item);
            SetStatus($"ðŸ—‘ï¸ Item removed ({_loanItems.Count} items remaining)", Brushes.Orange);
        }
    }

    /// <summary>Build a Loan object from the Loan form fields.</summary>
    private Loan? BuildLoanFromForm()
    {
        if (string.IsNullOrWhiteSpace(LoanCustName.Text) ||
            string.IsNullOrWhiteSpace(LoanPhone.Text) ||
            string.IsNullOrWhiteSpace(LoanAddress.Text) ||
            string.IsNullOrWhiteSpace(LoanGovId.Text) ||
            _loanItems.Count == 0 ||
            !decimal.TryParse(LoanPrincipal.Text, out var principal))
        {
            SetStatus("âŒ Please fill all required fields and add at least one item", Brushes.Red);
            return null;
        }

        decimal.TryParse(LoanInterest.Text, out var interest);
        var startDate = LoanStartDatePicker.SelectedDate?.ToString("yyyy-MM-dd")
            ?? DateTime.Now.ToString("yyyy-MM-dd");
        var idType = (LoanIdType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "AADHAAR";

        var items = _loanItems.ToList();

        return new Loan
        {
            Id = _excelService.GetNextId("Loans", "L-"),
            CustomerName = LoanCustName.Text.Trim(),
            CustomerPhone = LoanPhone.Text.Trim(),
            CustomerAddress = LoanAddress.Text.Trim(),
            GovId = LoanGovId.Text.Trim(),
            GovIdType = idType,
            MetalType = string.Join(" | ", items.Select(i => i.MetalType)),
            ProductDescription = string.Join(" | ", items.Select(i => i.ProductDescription)),
            Weight = items.Sum(i => i.Weight),
            Purity = string.Join(" | ", items.Select(i => i.Purity)),
            PrincipalAmount = principal,
            InterestRate = interest,
            StartDate = startDate,
            TotalRepaid = 0,
            Status = "ACTIVE",
            Items = items
        };
    }

    /// <summary>Show a preview of the invoice without saving (DRAFT).</summary>
    private void DummyInvoice_Click(object sender, RoutedEventArgs e)
    {
        var invoice = BuildInvoiceFromForm();
        if (invoice == null) return;

        invoice.Id = "DRAFT";
        var doc = BuildInvoiceDocument(invoice);

        ShowDocumentPreview(doc, "Dummy Invoice Preview â€” DRAFT");
        SetStatus("ðŸ“‹ Dummy invoice preview shown (not saved).", Brushes.Cyan);
    }

    /// <summary>Preview & Print a Loan receipt.</summary>
    private void PreviewLoanInvoice_Click(object sender, RoutedEventArgs e)
    {
        var loan = BuildLoanFromForm();
        if (loan == null) return;

        var doc = BuildLoanDocument(loan, isDraft: false);

        var pd = new PrintDialog();
        if (pd.ShowDialog() == true)
        {
            var paginator = ((IDocumentPaginatorSource)doc).DocumentPaginator;
            pd.PrintDocument(paginator, $"Loan Receipt {loan.Id}");
            SetStatus($"ðŸ–¨ï¸ Loan receipt {loan.Id} printed.", Brushes.Cyan);
        }
        else
        {
            SetStatus($"ðŸ–¨ï¸ Loan receipt {loan.Id} previewed (not printed).", Brushes.Cyan);
        }
    }

    /// <summary>Show a preview of the loan receipt without saving (DRAFT).</summary>
    private void DummyLoanInvoice_Click(object sender, RoutedEventArgs e)
    {
        var loan = BuildLoanFromForm();
        if (loan == null) return;

        loan.Id = "DRAFT";
        var doc = BuildLoanDocument(loan, isDraft: true);

        ShowDocumentPreview(doc, "Dummy Loan Receipt Preview â€” DRAFT");
        SetStatus("ðŸ“‹ Dummy loan receipt preview shown (not saved).", Brushes.Cyan);
    }

    /// <summary>Build a professional Loan Receipt FlowDocument (Half-A4 with Hindi).</summary>
    private static FlowDocument BuildLoanDocument(Loan loan, bool isDraft)
    {
        var doc = new FlowDocument
        {
            // Half-A4: 210mm Ã— 148.5mm = 794 Ã— 561 px at 96 DPI
            PageWidth = 794,
            PageHeight = 561,
            PagePadding = new Thickness(24, 16, 24, 12),
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 8,
            ColumnWidth = 999,
            IsColumnWidthFlexible = false
        };

        var hindiFont = new FontFamily("Nirmala UI");
        var borderBrush = Brushes.Black;
        var goldColor = new SolidColorBrush(Color.FromRgb(139, 101, 8));
        var headerBg = new SolidColorBrush(Color.FromRgb(245, 235, 210));
        var lightBg = new SolidColorBrush(Color.FromRgb(255, 252, 245));
        var accentBg = new SolidColorBrush(Color.FromRgb(250, 245, 225));
        var darkHeaderBg = new SolidColorBrush(Color.FromRgb(85, 65, 20));

        // â•â•â• 1. TITLE BAR â•â•â•
        var titleTable = new Table { CellSpacing = 0 };
        titleTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var titleGroup = new TableRowGroup();
        var titleRow = new TableRow { Background = darkHeaderBg };
        var titleText = isDraft ? "GOLD LOAN RECEIPT â€” DRAFT" : "GOLD LOAN RECEIPT";
        var titleP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        titleP.Inlines.Add(new Run(titleText) { FontSize = 11, FontWeight = FontWeights.Bold, Foreground = Brushes.White });
        titleP.Inlines.Add(new Run("  |  गिरवी रसीद") { FontSize = 10, FontWeight = FontWeights.Bold, Foreground = Brushes.Gold, FontFamily = hindiFont });
        var titleCell = new TableCell(titleP);
        titleCell.BorderBrush = borderBrush;
        titleCell.BorderThickness = new Thickness(1);
        titleCell.Padding = new Thickness(0, 2, 0, 2);
        titleRow.Cells.Add(titleCell);
        titleGroup.Rows.Add(titleRow);
        titleTable.RowGroups.Add(titleGroup);
        doc.Blocks.Add(titleTable);

        // â•â•â• 2. SHOP HEADER (compact) â•â•â•
        var shopTable = new Table { CellSpacing = 0 };
        shopTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var shopGroup = new TableRowGroup();
        var shopRow = new TableRow { Background = accentBg };
        var shopCell = new TableCell();
        shopCell.BorderBrush = borderBrush;
        shopCell.BorderThickness = new Thickness(1, 0, 1, 0);
        shopCell.Padding = new Thickness(4, 4, 4, 4);

        var shopP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0) };
        shopP.Inlines.Add(new Run("PHOOLCHANDRA SARAF JEWELLERS") { FontSize = 16, FontWeight = FontWeights.Bold, Foreground = goldColor });
        shopP.Inlines.Add(new Run("  (फूलचन्द्र सर्राफ ज्वैलर्स)") { FontSize = 11, FontWeight = FontWeights.Bold, Foreground = goldColor, FontFamily = hindiFont });
        shopCell.Blocks.Add(shopP);


        var addrP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 1, 0, 0) };
        addrP.Inlines.Add(new Run("Bahadur Shah Nagar, Koraon-Prayagraj  |  Ph: 7518318070") { FontSize = 7, Foreground = Brushes.DimGray });
        shopCell.Blocks.Add(addrP);

        shopRow.Cells.Add(shopCell);
        shopGroup.Rows.Add(shopRow);
        shopTable.RowGroups.Add(shopGroup);
        doc.Blocks.Add(shopTable);

        // â•â•â• 3. META + CUSTOMER (side by side for compactness) â•â•â•
        var infoTable = new Table { CellSpacing = 0 };
        infoTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        infoTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var infoGroup = new TableRowGroup();
        var infoRow = new TableRow();

        // Left: Loan details
        var loanCell = new TableCell();
        loanCell.BorderBrush = borderBrush;
        loanCell.BorderThickness = new Thickness(1, 0, 0.5, 1);
        loanCell.Padding = new Thickness(6, 3, 6, 3);
        var loanInfo = new Paragraph { FontSize = 8, LineHeight = 13, Margin = new Thickness(0) };
        loanInfo.Inlines.Add(new Run("Loan ID: ") { Foreground = Brushes.Gray });
        loanInfo.Inlines.Add(new Run(loan.Id) { FontWeight = FontWeights.Bold });
        loanInfo.Inlines.Add(new Run($"\nDate: {loan.StartDate}"));
        loanInfo.Inlines.Add(new Run($"\nPrincipal: â‚¹{loan.PrincipalAmount:N2}") { FontWeight = FontWeights.Bold });
        loanInfo.Inlines.Add(new Run($"\nInterest: {loan.InterestRate}%/month"));
        loanInfo.Inlines.Add(new Run($"\nMonthly Interest: â‚¹{loan.MonthlyInterest:N2}") { FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(180, 0, 0)) });
        var govLabel = !string.IsNullOrWhiteSpace(loan.GovIdType) ? loan.GovIdType : "Gov ID";
        loanInfo.Inlines.Add(new Run($"\n{govLabel}: {loan.GovId}"));
        loanCell.Blocks.Add(loanInfo);
        infoRow.Cells.Add(loanCell);

        // Right: Customer details
        var custCell = new TableCell();
        custCell.BorderBrush = borderBrush;
        custCell.BorderThickness = new Thickness(0.5, 0, 1, 1);
        custCell.Padding = new Thickness(6, 3, 6, 3);
        var custInfo = new Paragraph { FontSize = 8, LineHeight = 13, Margin = new Thickness(0) };
        custInfo.Inlines.Add(new Run("Customer / ग्राहक:") { Foreground = Brushes.Gray, FontFamily = hindiFont });
        custInfo.Inlines.Add(new Run($"\n{loan.CustomerName}") { FontWeight = FontWeights.Bold });
        custInfo.Inlines.Add(new Run($"\nPhone: {loan.CustomerPhone}"));
        custInfo.Inlines.Add(new Run($"\nAddress: {loan.CustomerAddress}"));
        custInfo.Inlines.Add(new Run($"\nStatus: {loan.Status}") { FontWeight = FontWeights.SemiBold });
        custCell.Blocks.Add(custInfo);
        infoRow.Cells.Add(custCell);

        infoGroup.Rows.Add(infoRow);
        infoTable.RowGroups.Add(infoGroup);
        doc.Blocks.Add(infoTable);

        // â•â•â• 4. PLEDGED ITEM TABLE (NO PURITY) â•â•â•
        var itemTable = new Table { CellSpacing = 0 };
        // 5 columns â€” no Purity. Sum = 746px (794 - 48 padding)
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(30) });       // Sr
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(346) });      // Product
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });       // Metal
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(80) });       // Weight
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(120) });      // Principal

        var itemGroup = new TableRowGroup();

        // Header
        var hdr = new TableRow { Background = headerBg };
        AddItemCell(hdr, "Sr", FontWeights.Bold, TextAlignment.Center, 7);
        AddItemCell(hdr, "Product Description / विवरण", FontWeights.Bold, TextAlignment.Left, 7);
        AddItemCell(hdr, "Metal", FontWeights.Bold, TextAlignment.Center, 7);
        AddItemCell(hdr, "Weight (g)", FontWeights.Bold, TextAlignment.Right, 7);
        AddItemCell(hdr, "Principal (â‚¹)", FontWeights.Bold, TextAlignment.Right, 7);
        itemGroup.Rows.Add(hdr);

        // Data rows (one per item â€” NO purity column)
        var loanItemsList = loan.Items != null && loan.Items.Count > 0
            ? loan.Items
            : new List<LoanItem> { new LoanItem {
                ProductDescription = loan.ProductDescription, MetalType = loan.MetalType,
                Purity = loan.Purity, Weight = loan.Weight } };

        for (int i = 0; i < loanItemsList.Count; i++)
        {
            var li = loanItemsList[i];
            var bg = (i % 2 == 0) ? lightBg : accentBg;
            var dRow = new TableRow { Background = bg };
            AddItemCell(dRow, $"{i + 1}", FontWeights.Normal, TextAlignment.Center, 7);
            AddItemCell(dRow, li.ProductDescription, FontWeights.Normal, TextAlignment.Left, 7);
            AddItemCell(dRow, li.MetalType, FontWeights.Normal, TextAlignment.Center, 7);
            AddItemCell(dRow, $"{li.Weight:N3}", FontWeights.Normal, TextAlignment.Right, 7);
            AddItemCell(dRow, i == 0 ? $"{loan.PrincipalAmount:N2}" : "", FontWeights.Normal, TextAlignment.Right, 7);
            itemGroup.Rows.Add(dRow);
        }

        // Interest summary row (4 col span, no purity)
        var intRow = new TableRow();
        intRow.Cells.Add(MakeSpanCell($"Interest @ {loan.InterestRate}% per month (à¤®à¤¾à¤¸à¤¿à¤• à¤¬à¥à¤¯à¤¾à¤œ)", 4, TextAlignment.Right, FontWeights.Normal, 7));
        AddItemCell(intRow, $"â‚¹{loan.MonthlyInterest:N2}/month", FontWeights.Bold, TextAlignment.Right, 7);
        itemGroup.Rows.Add(intRow);

        // Principal total row
        var totalRow = new TableRow { Background = darkHeaderBg };
        var totalLabel = new TableCell(new Paragraph(new Run("Total Principal / à¤•à¥à¤² à¤®à¥‚à¤²à¤§à¤¨"))
        {
            FontSize = 9, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(2, 2, 2, 2)
        }) { ColumnSpan = 4 };
        totalLabel.BorderBrush = borderBrush;
        totalLabel.BorderThickness = new Thickness(0.5);
        totalRow.Cells.Add(totalLabel);

        var totalVal = new TableCell(new Paragraph(new Run($"â‚¹ {loan.PrincipalAmount:N2}"))
        {
            FontSize = 10, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(2, 2, 2, 2)
        });
        totalVal.BorderBrush = borderBrush;
        totalVal.BorderThickness = new Thickness(0.5);
        totalRow.Cells.Add(totalVal);
        itemGroup.Rows.Add(totalRow);

        itemTable.RowGroups.Add(itemGroup);
        SetTableBorder(itemTable);
        doc.Blocks.Add(itemTable);

        // â•â•â• 5. AMOUNT IN WORDS â•â•â•
        var wordsP = new Paragraph { FontSize = 7, Margin = new Thickness(0, 2, 0, 2) };
        wordsP.Inlines.Add(new Run("In Words / शब्दों में: ") { FontWeight = FontWeights.Bold, FontFamily = hindiFont });
        wordsP.Inlines.Add(new Run(NumberToWords((long)Math.Round(loan.PrincipalAmount)) + " Rupees Only") { FontStyle = FontStyles.Italic, FontWeight = FontWeights.SemiBold });
        doc.Blocks.Add(wordsP);

        // â•â•â• 6. TERMS & CONDITIONS (36 months) + SIGNATURE (side by side) â•â•â•
        var footTable = new Table { CellSpacing = 0 };
        footTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        footTable.Columns.Add(new TableColumn { Width = new GridLength(220) });
        var footGroup = new TableRowGroup();
        var footRow = new TableRow();

        // Left: Terms
        var termsCell = new TableCell();
        termsCell.BorderBrush = borderBrush;
        termsCell.BorderThickness = new Thickness(1, 1, 0.5, 1);
        termsCell.Padding = new Thickness(4, 2, 4, 2);

        var termsHeader = new Paragraph { Margin = new Thickness(0, 0, 0, 2) };
        termsHeader.Inlines.Add(new Run("Terms & Conditions") { FontSize = 7, FontWeight = FontWeights.Bold });
        termsHeader.Inlines.Add(new Run(" / à¤¨à¤¿à¤¯à¤® à¤”à¤° à¤¶à¤°à¥à¤¤à¥‡à¤‚") { FontSize = 7, FontWeight = FontWeights.Bold, FontFamily = hindiFont });
        termsCell.Blocks.Add(termsHeader);

        var terms = new List { FontSize = 6.5, Foreground = Brushes.DimGray, MarkerStyle = TextMarkerStyle.Decimal, Padding = new Thickness(12, 0, 0, 0) };
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Loan duration is 36 months from the date of pledge. / ऋण अवधि गिरवी रखने की तिथि से 36 माह है।")) { FontFamily = hindiFont }));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Interest is payable monthly. Non-payment may incur additional charges."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("If not redeemed within 36 months, pledged article(s) may be auctioned."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("This receipt must be produced at the time of redemption."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Shop is not liable for natural wear or damage during storage."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Disputes subject to local jurisdiction."))));
        termsCell.Blocks.Add(terms);
        footRow.Cells.Add(termsCell);

        // Right: Signatures
        var sigCell = new TableCell();
        sigCell.BorderBrush = borderBrush;
        sigCell.BorderThickness = new Thickness(0.5, 1, 1, 1);
        sigCell.Padding = new Thickness(4, 2, 4, 2);

        sigCell.Blocks.Add(new Paragraph(new Run("Pledger's Signature / गिरवीकर्ता")) { FontSize = 7, FontWeight = FontWeights.Bold, FontFamily = hindiFont, Margin = new Thickness(0, 0, 0, 6) });
        sigCell.Blocks.Add(new Paragraph(new Run("__________________________")) { FontSize = 7, Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 0, 10) });

        var forP = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0, 0, 0, 2) };
        forP.Inlines.Add(new Run("For, ") { FontSize = 7, Foreground = Brushes.Gray });
        forP.Inlines.Add(new Run("PHOOLCHANDRA SARAF JEWELLERS") { FontSize = 8, FontWeight = FontWeights.Bold });
        sigCell.Blocks.Add(forP);
        sigCell.Blocks.Add(new Paragraph(new Run("(Authorised Signatory)")) { FontSize = 6, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Right });
        footRow.Cells.Add(sigCell);

        footGroup.Rows.Add(footRow);
        footTable.RowGroups.Add(footGroup);
        doc.Blocks.Add(footTable);

        // â•â•â• 7. FOOTER â•â•â•
        var footerText = isDraft ? "âš  DRAFT â€” This is NOT an official loan receipt." : "This is a computer generated loan receipt.";
        doc.Blocks.Add(new Paragraph(new Run(footerText))
        {
            FontSize = 6, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 3, 0, 0)
        });

        return doc;
    }

    /// <summary>Show a FlowDocument in a preview window (no printing).</summary>
    private void ShowDocumentPreview(FlowDocument doc, string title)
    {
        var reader = new FlowDocumentReader
        {
            Document = doc,
            ViewingMode = FlowDocumentReaderViewingMode.Page,
            Background = Brushes.White
        };

        var previewWindow = new Window
        {
            Title = title,
            Width = 850,
            Height = 1000,
            WindowStartupLocation = WindowStartupLocation.CenterScreen,
            Content = reader,
            Owner = this
        };
        previewWindow.ShowDialog();
    }

    private void OpenExcel_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (File.Exists(_excelPath))
            {
                Process.Start(new ProcessStartInfo(_excelPath) { UseShellExecute = true });
                SetStatus("ðŸ“‚ Opened Excel file", Brushes.Cyan);
            }
            else
            {
                SetStatus("âš ï¸ Excel file not found â€” save some data first!", Brushes.Orange);
            }
        }
        catch (Exception ex)
        {
            SetStatus($"âŒ Cannot open Excel: {ex.Message}", Brushes.Red);
        }
    }

    private void Refresh_Click(object sender, RoutedEventArgs e) => LoadAllData();

    private void Tab_Changed(object sender, SelectionChangedEventArgs e) { }

    // â”€â”€â”€ HELPERS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

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
