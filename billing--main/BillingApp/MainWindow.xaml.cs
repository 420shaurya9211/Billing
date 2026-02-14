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

        var billType = (InvBillType.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "PAKKA";
        var subTotal = (weight * rate) + making - discount;

        // KACHA bills have no GST
        decimal cgst = 0, sgst = 0, gstAmount = 0, total;
        if (billType == "PAKKA")
        {
            cgst = Math.Round(subTotal * 0.015m, 2);
            sgst = cgst;
            gstAmount = cgst + sgst;
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

        return new Invoice
        {
            Id = _excelService.GetNextId("Invoices", "INV-"),
            CustomerId = InvCustId.Text.Trim(),
            CustomerAddress = InvAddress.Text.Trim(),
            Date = DateTime.Now.ToString("yyyy-MM-dd"),
            BillType = billType,
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

    /// <summary>Build a professional Gold Tax Invoice FlowDocument matching jewellery invoice standards.</summary>
    private static FlowDocument BuildInvoiceDocument(Invoice inv)
    {
        var doc = new FlowDocument
        {
            // A4 paper: 210mm × 297mm = 794 × 1123 px at 96 DPI
            PageWidth = 794,
            PageHeight = 1123,
            PagePadding = new Thickness(40, 30, 40, 30),
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 10,
            ColumnWidth = 999,                // force single-column layout
            IsColumnWidthFlexible = false
        };

        var borderBrush = Brushes.Black;
        var goldColor = new SolidColorBrush(Color.FromRgb(139, 101, 8));
        var headerBg = new SolidColorBrush(Color.FromRgb(245, 235, 210));   // warm cream
        var lightBg = new SolidColorBrush(Color.FromRgb(255, 252, 245));
        var accentBg = new SolidColorBrush(Color.FromRgb(250, 245, 225));
        var darkHeaderBg = new SolidColorBrush(Color.FromRgb(85, 65, 20));

        // ═══════════════════════════════════════════════════════════════
        // 1. TITLE BAR - "GOLD TAX INVOICE" (dark bar)
        // ═══════════════════════════════════════════════════════════════
        var titleTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        titleTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var titleGroup = new TableRowGroup();
        var titleRow = new TableRow { Background = darkHeaderBg };
        var titleCell = new TableCell(new Paragraph(new Run("GOLD TAX INVOICE"))
        {
            FontSize = 13, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 4, 0, 4), Foreground = Brushes.White
        });
        titleCell.BorderBrush = borderBrush;
        titleCell.BorderThickness = new Thickness(1);
        titleRow.Cells.Add(titleCell);
        titleGroup.Rows.Add(titleRow);
        titleTable.RowGroups.Add(titleGroup);
        doc.Blocks.Add(titleTable);

        // ═══════════════════════════════════════════════════════════════
        // 2. SHOP HEADER BLOCK (bordered)
        // ═══════════════════════════════════════════════════════════════
        var shopTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        shopTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var shopGroup = new TableRowGroup();

        // Shop Name Row
        var shopNameRow = new TableRow { Background = accentBg };
        var shopNameCell = new TableCell();
        shopNameCell.BorderBrush = borderBrush;
        shopNameCell.BorderThickness = new Thickness(1, 0, 1, 0);

        var shopNameP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 8, 0, 0) };
        shopNameP.Inlines.Add(new Run("PHOOL CHANDRA SARAF") { FontSize = 24, FontWeight = FontWeights.Bold, Foreground = goldColor });
        shopNameCell.Blocks.Add(shopNameP);

        var shopSubP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        shopSubP.Inlines.Add(new Run("ASHISH JEWELLERS") { FontSize = 16, FontWeight = FontWeights.Bold, Foreground = Brushes.Black });
        shopNameCell.Blocks.Add(shopSubP);

        // Tagline
        var tagP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        tagP.Inlines.Add(new Run("91.6 Hallmark Showroom") { FontSize = 9, FontWeight = FontWeights.SemiBold, Foreground = Brushes.DarkGoldenrod });
        tagP.Inlines.Add(new Run("     |     ") { FontSize = 9, Foreground = Brushes.Gray });
        tagP.Inlines.Add(new Run("Gold & Silver Ornament Traders") { FontSize = 9, FontWeight = FontWeights.SemiBold, Foreground = Brushes.DarkGoldenrod });
        shopNameCell.Blocks.Add(tagP);

        // Address
        var addrP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 4, 0, 8) };
        addrP.Inlines.Add(new Run("Koraon, Allahabad, Uttar Pradesh - 212306  |  Ph: 7985494707") { FontSize = 9, Foreground = Brushes.DimGray });
        shopNameCell.Blocks.Add(addrP);

        shopNameRow.Cells.Add(shopNameCell);
        shopGroup.Rows.Add(shopNameRow);
        shopTable.RowGroups.Add(shopGroup);
        doc.Blocks.Add(shopTable);

        // ═══════════════════════════════════════════════════════════════
        // 3. INVOICE META TABLE (GSTIN, Invoice No, Date, Bill Type, Status)
        // ═══════════════════════════════════════════════════════════════
        var metaTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        metaTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        metaTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        metaTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var metaGroup = new TableRowGroup();

        // Row 1: GSTIN | Invoice No | Dated
        var metaRow1 = new TableRow();
        AddLabelValueCell(metaRow1, "GSTIN", "N/A");
        AddLabelValueCell(metaRow1, "Invoice No.", inv.Id);
        AddLabelValueCell(metaRow1, "Dated", inv.Date);
        metaGroup.Rows.Add(metaRow1);

        // Row 2: Bill Type | Status | Payment
        var metaRow2 = new TableRow();
        AddLabelValueCell(metaRow2, "Bill Type", inv.BillType);
        AddLabelValueCell(metaRow2, "Status", inv.Status);
        AddLabelValueCell(metaRow2, "Pay Mode", inv.Status == "PAID" ? "Cash" : "Pending");
        metaGroup.Rows.Add(metaRow2);

        metaTable.RowGroups.Add(metaGroup);
        SetTableBorder(metaTable);
        doc.Blocks.Add(metaTable);

        // ═══════════════════════════════════════════════════════════════
        // 4. CUSTOMER DETAILS (Billed To) — bordered box
        // ═══════════════════════════════════════════════════════════════
        var custTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        custTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var custGroup = new TableRowGroup();
        var custRow = new TableRow();
        var custCell = new TableCell();
        custCell.BorderBrush = borderBrush;
        custCell.BorderThickness = new Thickness(1, 0, 1, 1);
        custCell.Padding = new Thickness(8, 4, 8, 6);

        var custHeader = new Paragraph(new Run("Details of Receiver (Billed To)"))
        {
            FontSize = 10, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 4),
            TextDecorations = TextDecorations.Underline
        };
        custCell.Blocks.Add(custHeader);

        var custInfo = new Paragraph { Margin = new Thickness(0), FontSize = 10, LineHeight = 18 };
        custInfo.Inlines.Add(new Run("Customer: ") { Foreground = Brushes.Gray });
        custInfo.Inlines.Add(new Run(inv.CustomerId) { FontWeight = FontWeights.Bold });
        if (!string.IsNullOrWhiteSpace(inv.CustomerAddress))
        {
            custInfo.Inlines.Add(new Run("\nAddress: ") { Foreground = Brushes.Gray });
            custInfo.Inlines.Add(new Run(inv.CustomerAddress));
        }
        custCell.Blocks.Add(custInfo);

        custRow.Cells.Add(custCell);
        custGroup.Rows.Add(custRow);
        custTable.RowGroups.Add(custGroup);
        doc.Blocks.Add(custTable);

        // ═══════════════════════════════════════════════════════════════
        // 5. UNIFIED ITEM + GST + GRAND TOTAL TABLE
        //    (all in one table so columns align perfectly)
        // ═══════════════════════════════════════════════════════════════
        var itemTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        // 8 columns — all fixed widths (Star doesn't render reliably in FlowDocument tables)
        // Available width: 794 - 80 (margins) = 714px
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(30) });      // Col 0: Sr no
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(200) });     // Col 1: Product Name & HSN
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(45) });      // Col 2: Purity
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(62) });      // Col 3: Gross Wt
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(62) });      // Col 4: Net Wt
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(65) });      // Col 5: Rate/Unit
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });      // Col 6: Making/Unit
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(96) });      // Col 7: Amount

        var itemGroup = new TableRowGroup();

        // ── Header row ──
        var hdr = new TableRow { Background = headerBg };
        AddItemCell(hdr, "Sr no", FontWeights.Bold, TextAlignment.Center, 8);
        AddItemCell(hdr, "Product Name & HSN", FontWeights.Bold, TextAlignment.Left, 8);
        AddItemCell(hdr, "Purity", FontWeights.Bold, TextAlignment.Center, 8);
        AddItemCell(hdr, "Gross Wt", FontWeights.Bold, TextAlignment.Right, 8);
        AddItemCell(hdr, "Net Wt", FontWeights.Bold, TextAlignment.Right, 8);
        AddItemCell(hdr, "Rate/Unit", FontWeights.Bold, TextAlignment.Right, 8);
        AddItemCell(hdr, "Making/Unit", FontWeights.Bold, TextAlignment.Right, 8);
        var amountLabel = inv.BillType == "PAKKA" ? "Taxable Amount" : "Total Amount";
        AddItemCell(hdr, amountLabel, FontWeights.Bold, TextAlignment.Right, 8);
        itemGroup.Rows.Add(hdr);

        // ── Data row ──
        var dRow = new TableRow { Background = lightBg };
        AddItemCell(dRow, "1", FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(dRow, $"{inv.ItemDescription} ({inv.Metal})", FontWeights.Normal, TextAlignment.Left, 9);
        AddItemCell(dRow, inv.Purity, FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(dRow, $"{inv.Weight:N3}", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(dRow, $"{inv.Weight:N3}", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(dRow, $"{inv.RatePerGram:N0}", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(dRow, $"{inv.MakingCharges:N1}", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(dRow, $"{inv.SubTotal:N2}", FontWeights.Normal, TextAlignment.Right, 9);
        itemGroup.Rows.Add(dRow);

        // ── Total Pcs row ──
        var tRow = new TableRow { Background = headerBg };
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(tRow, "Total Pcs: 1", FontWeights.Bold, TextAlignment.Left, 9);
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(tRow, $"{inv.Weight:N3}", FontWeights.Bold, TextAlignment.Right, 9);
        AddItemCell(tRow, $"{inv.Weight:N3}", FontWeights.Bold, TextAlignment.Right, 9);
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(tRow, "", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(tRow, $"{inv.SubTotal:N2}", FontWeights.Bold, TextAlignment.Right, 9);
        itemGroup.Rows.Add(tRow);

        // ── Discount row (if any) — spans cols 0-6, amount in col 7 ──
        if (inv.Discount > 0)
        {
            var discRow = new TableRow();
            var discLbl = MakeSpanCell($"Less Discount:", 7, TextAlignment.Right, FontWeights.Normal, 9);
            discRow.Cells.Add(discLbl);
            AddItemCell(discRow, $"-{inv.Discount:N2}", FontWeights.Bold, TextAlignment.Right, 9);
            itemGroup.Rows.Add(discRow);
        }

        // ── CGST / SGST rows — only for PAKKA bills ──
        if (inv.BillType == "PAKKA")
        {
            var cgstRow = new TableRow();
            cgstRow.Cells.Add(MakeSpanCell($"CGST @ {inv.CgstRate}%", 7, TextAlignment.Right, FontWeights.Normal, 9));
            AddItemCell(cgstRow, $"{(inv.GstAmount / 2):N2}", FontWeights.Bold, TextAlignment.Right, 9);
            itemGroup.Rows.Add(cgstRow);

            var sgstRow = new TableRow();
            sgstRow.Cells.Add(MakeSpanCell($"SGST @ {inv.SgstRate}%", 7, TextAlignment.Right, FontWeights.Normal, 9));
            AddItemCell(sgstRow, $"{(inv.GstAmount / 2):N2}", FontWeights.Bold, TextAlignment.Right, 9);
            itemGroup.Rows.Add(sgstRow);
        }

        // ── Return row (if any) ──
        if (inv.ReturnAmount > 0)
        {
            var retLabel = inv.ReturnWeight > 0 ? $"Return ({inv.ReturnWeight:N3}g)" : "Return Adjustment";
            var retRow = new TableRow();
            retRow.Cells.Add(MakeSpanCell(retLabel, 7, TextAlignment.Right, FontWeights.Normal, 9));
            AddItemCell(retRow, $"-{inv.ReturnAmount:N2}", FontWeights.Bold, TextAlignment.Right, 9);
            itemGroup.Rows.Add(retRow);
        }

        // ── Grand Total row — "In Words" on left, "Grand Total:" + amount on right ──
        var grandRow = new TableRow { Background = darkHeaderBg };
        var wordsText = "In Words: " + NumberToWords((long)Math.Round(inv.NetAmount)) + " Rupees Only";
        var wordsCellGrand = new TableCell(new Paragraph(new Run(wordsText))
        {
            FontSize = 8, FontWeight = FontWeights.SemiBold, FontStyle = FontStyles.Italic,
            Foreground = Brushes.White, Margin = new Thickness(4, 3, 4, 3)
        }) { ColumnSpan = 6 };
        wordsCellGrand.BorderBrush = borderBrush;
        wordsCellGrand.BorderThickness = new Thickness(0.5);
        grandRow.Cells.Add(wordsCellGrand);

        var gtLabelCell = new TableCell(new Paragraph(new Run("Grand Total :"))
        {
            FontSize = 11, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(4, 3, 4, 3)
        });
        gtLabelCell.BorderBrush = borderBrush;
        gtLabelCell.BorderThickness = new Thickness(0.5);
        grandRow.Cells.Add(gtLabelCell);

        var gtValCell = new TableCell(new Paragraph(new Run($"₹ {inv.NetAmount:N2}"))
        {
            FontSize = 12, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(4, 3, 4, 3)
        });
        gtValCell.BorderBrush = borderBrush;
        gtValCell.BorderThickness = new Thickness(0.5);
        grandRow.Cells.Add(gtValCell);
        itemGroup.Rows.Add(grandRow);

        itemTable.RowGroups.Add(itemGroup);
        SetTableBorder(itemTable);
        doc.Blocks.Add(itemTable);

        // ═══════════════════════════════════════════════════════════════
        // 8. TERMS & CONDITIONS + BANK DETAILS (side by side)
        // ═══════════════════════════════════════════════════════════════
        var footerTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        footerTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        footerTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var footerGroup = new TableRowGroup();
        var footerRow = new TableRow();

        // Left: Terms
        var termsCell = new TableCell();
        termsCell.BorderBrush = borderBrush;
        termsCell.BorderThickness = new Thickness(1, 0, 0.5, 1);
        termsCell.Padding = new Thickness(6, 4, 6, 4);

        termsCell.Blocks.Add(new Paragraph(new Run("Terms & Conditions:"))
        {
            FontSize = 8, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 2)
        });

        var terms = new List { MarkerStyle = TextMarkerStyle.Decimal, FontSize = 7, Foreground = Brushes.DimGray, Margin = new Thickness(12, 0, 0, 0) };
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Check the goods and weight while buying."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Bill must be brought for return/exchange."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Labour and Taxes are non-refundable."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("KDM Jewellery: 15% depreciation."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("916 Hallmark: 10% depreciation."))));
        termsCell.Blocks.Add(terms);
        footerRow.Cells.Add(termsCell);

        // Right: Bank Details
        var bankCell = new TableCell();
        bankCell.BorderBrush = borderBrush;
        bankCell.BorderThickness = new Thickness(0.5, 0, 1, 1);
        bankCell.Padding = new Thickness(6, 4, 6, 4);

        bankCell.Blocks.Add(new Paragraph(new Run("Bank Details:"))
        {
            FontSize = 8, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 2)
        });

        var bankInfo = new Paragraph { FontSize = 8, Foreground = Brushes.DimGray, LineHeight = 14 };
        bankInfo.Inlines.Add(new Run("Bank Name: ") { Foreground = Brushes.Gray });
        bankInfo.Inlines.Add(new Run("____________\n") { FontWeight = FontWeights.SemiBold });
        bankInfo.Inlines.Add(new Run("A/c No: ") { Foreground = Brushes.Gray });
        bankInfo.Inlines.Add(new Run("____________\n") { FontWeight = FontWeights.SemiBold });
        bankInfo.Inlines.Add(new Run("IFSC Code: ") { Foreground = Brushes.Gray });
        bankInfo.Inlines.Add(new Run("____________") { FontWeight = FontWeights.SemiBold });
        bankCell.Blocks.Add(bankInfo);
        footerRow.Cells.Add(bankCell);

        footerGroup.Rows.Add(footerRow);
        footerTable.RowGroups.Add(footerGroup);
        doc.Blocks.Add(footerTable);

        // ═══════════════════════════════════════════════════════════════
        // 9. SIGNATURE AREA (bordered, side by side)
        // ═══════════════════════════════════════════════════════════════
        var sigTable = new Table { CellSpacing = 0, Margin = new Thickness(0, 0, 0, 0) };
        sigTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        sigTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var sigGroup = new TableRowGroup();
        var sigRow = new TableRow();

        // Left: Customer Signature
        var leftSig = new TableCell();
        leftSig.BorderBrush = borderBrush;
        leftSig.BorderThickness = new Thickness(1, 0, 0.5, 1);
        leftSig.Padding = new Thickness(8, 4, 8, 4);
        var leftP1 = new Paragraph(new Run("Customer Signature"))
        {
            FontSize = 9, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 30)
        };
        leftSig.Blocks.Add(leftP1);
        var leftP2 = new Paragraph(new Run("__________________________"))
        {
            FontSize = 8, Foreground = Brushes.Gray, Margin = new Thickness(0)
        };
        leftSig.Blocks.Add(leftP2);
        sigRow.Cells.Add(leftSig);

        // Right: For Shop
        var rightSig = new TableCell();
        rightSig.BorderBrush = borderBrush;
        rightSig.BorderThickness = new Thickness(0.5, 0, 1, 1);
        rightSig.Padding = new Thickness(8, 4, 8, 4);

        var rightP1 = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0, 0, 0, 4) };
        rightP1.Inlines.Add(new Run("For, ") { FontSize = 9, Foreground = Brushes.Gray });
        rightP1.Inlines.Add(new Run("ASHISH JEWELLERS") { FontSize = 10, FontWeight = FontWeights.Bold });
        rightSig.Blocks.Add(rightP1);

        var rightP2 = new Paragraph(new Run(""))
        {
            FontSize = 8, Margin = new Thickness(0, 0, 0, 16)
        };
        rightSig.Blocks.Add(rightP2);

        var rightP3 = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0) };
        rightP3.Inlines.Add(new Run("(Authorised Signatory)") { FontSize = 8, Foreground = Brushes.Gray });
        rightSig.Blocks.Add(rightP3);
        sigRow.Cells.Add(rightSig);

        sigGroup.Rows.Add(sigRow);
        sigTable.RowGroups.Add(sigGroup);
        doc.Blocks.Add(sigTable);

        // ═══════════════════════════════════════════════════════════════
        // 10. FOOTER NOTE
        // ═══════════════════════════════════════════════════════════════
        doc.Blocks.Add(new Paragraph(new Run("This is a computer generated invoice."))
        {
            FontSize = 7, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 6, 0, 0)
        });

        return doc;
    }

    // ─── Invoice Document Helpers ──────────────────────────────────────

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

    // ─── LOAN INVOICE & DUMMY INVOICE ─────────────────────────────────

    /// <summary>Build a Loan object from the Loan form fields.</summary>
    private Loan? BuildLoanFromForm()
    {
        if (string.IsNullOrWhiteSpace(LoanCustName.Text) ||
            string.IsNullOrWhiteSpace(LoanPhone.Text) ||
            string.IsNullOrWhiteSpace(LoanAddress.Text) ||
            string.IsNullOrWhiteSpace(LoanGovId.Text) ||
            string.IsNullOrWhiteSpace(LoanProduct.Text) ||
            !decimal.TryParse(LoanWeight.Text, out var weight) ||
            !decimal.TryParse(LoanPrincipal.Text, out var principal))
        {
            SetStatus("❌ Please fill all required loan fields", Brushes.Red);
            return null;
        }

        decimal.TryParse(LoanInterest.Text, out var interest);
        var startDate = string.IsNullOrWhiteSpace(LoanStartDate.Text)
            ? DateTime.Now.ToString("yyyy-MM-dd")
            : LoanStartDate.Text.Trim();

        return new Loan
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
    }

    /// <summary>Show a preview of the invoice without saving (DRAFT).</summary>
    private void DummyInvoice_Click(object sender, RoutedEventArgs e)
    {
        var invoice = BuildInvoiceFromForm();
        if (invoice == null) return;

        invoice.Id = "DRAFT";
        var doc = BuildInvoiceDocument(invoice);

        ShowDocumentPreview(doc, "Dummy Invoice Preview — DRAFT");
        SetStatus("📋 Dummy invoice preview shown (not saved).", Brushes.Cyan);
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
            SetStatus($"🖨️ Loan receipt {loan.Id} printed.", Brushes.Cyan);
        }
        else
        {
            SetStatus($"🖨️ Loan receipt {loan.Id} previewed (not printed).", Brushes.Cyan);
        }
    }

    /// <summary>Show a preview of the loan receipt without saving (DRAFT).</summary>
    private void DummyLoanInvoice_Click(object sender, RoutedEventArgs e)
    {
        var loan = BuildLoanFromForm();
        if (loan == null) return;

        loan.Id = "DRAFT";
        var doc = BuildLoanDocument(loan, isDraft: true);

        ShowDocumentPreview(doc, "Dummy Loan Receipt Preview — DRAFT");
        SetStatus("📋 Dummy loan receipt preview shown (not saved).", Brushes.Cyan);
    }

    /// <summary>Build a professional Loan Receipt FlowDocument.</summary>
    private static FlowDocument BuildLoanDocument(Loan loan, bool isDraft)
    {
        var doc = new FlowDocument
        {
            PageWidth = 794,
            PageHeight = 1123,
            PagePadding = new Thickness(40, 30, 40, 30),
            FontFamily = new FontFamily("Segoe UI"),
            FontSize = 10,
            ColumnWidth = 999,
            IsColumnWidthFlexible = false
        };

        var borderBrush = Brushes.Black;
        var goldColor = new SolidColorBrush(Color.FromRgb(139, 101, 8));
        var headerBg = new SolidColorBrush(Color.FromRgb(245, 235, 210));
        var lightBg = new SolidColorBrush(Color.FromRgb(255, 252, 245));
        var accentBg = new SolidColorBrush(Color.FromRgb(250, 245, 225));
        var darkHeaderBg = new SolidColorBrush(Color.FromRgb(85, 65, 20));

        // ═══ 1. TITLE BAR ═══
        var titleTable = new Table { CellSpacing = 0 };
        titleTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var titleGroup = new TableRowGroup();
        var titleRow = new TableRow { Background = darkHeaderBg };
        var titleText = isDraft ? "GOLD LOAN RECEIPT — DRAFT" : "GOLD LOAN RECEIPT";
        var titleCell = new TableCell(new Paragraph(new Run(titleText))
        {
            FontSize = 13, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 4, 0, 4), Foreground = Brushes.White
        });
        titleCell.BorderBrush = borderBrush;
        titleCell.BorderThickness = new Thickness(1);
        titleRow.Cells.Add(titleCell);
        titleGroup.Rows.Add(titleRow);
        titleTable.RowGroups.Add(titleGroup);
        doc.Blocks.Add(titleTable);

        // ═══ 2. SHOP HEADER ═══
        var shopTable = new Table { CellSpacing = 0 };
        shopTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var shopGroup = new TableRowGroup();
        var shopRow = new TableRow { Background = accentBg };
        var shopCell = new TableCell();
        shopCell.BorderBrush = borderBrush;
        shopCell.BorderThickness = new Thickness(1, 0, 1, 0);

        var shopP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 8, 0, 0) };
        shopP.Inlines.Add(new Run("PHOOL CHANDRA SARAF") { FontSize = 24, FontWeight = FontWeights.Bold, Foreground = goldColor });
        shopCell.Blocks.Add(shopP);

        var subP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        subP.Inlines.Add(new Run("ASHISH JEWELLERS") { FontSize = 16, FontWeight = FontWeights.Bold });
        shopCell.Blocks.Add(subP);

        var tagP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 2, 0, 0) };
        tagP.Inlines.Add(new Run("91.6 Hallmark Showroom") { FontSize = 9, FontWeight = FontWeights.SemiBold, Foreground = Brushes.DarkGoldenrod });
        tagP.Inlines.Add(new Run("     |     ") { FontSize = 9, Foreground = Brushes.Gray });
        tagP.Inlines.Add(new Run("Gold & Silver Ornament Traders") { FontSize = 9, FontWeight = FontWeights.SemiBold, Foreground = Brushes.DarkGoldenrod });
        shopCell.Blocks.Add(tagP);

        var addrP = new Paragraph { TextAlignment = TextAlignment.Center, Margin = new Thickness(0, 4, 0, 8) };
        addrP.Inlines.Add(new Run("Koraon, Allahabad, Uttar Pradesh - 212306  |  Ph: 7985494707") { FontSize = 9, Foreground = Brushes.DimGray });
        shopCell.Blocks.Add(addrP);

        shopRow.Cells.Add(shopCell);
        shopGroup.Rows.Add(shopRow);
        shopTable.RowGroups.Add(shopGroup);
        doc.Blocks.Add(shopTable);

        // ═══ 3. LOAN META TABLE ═══
        var metaTable = new Table { CellSpacing = 0 };
        metaTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        metaTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        metaTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var metaGroup = new TableRowGroup();

        var metaRow1 = new TableRow();
        AddLabelValueCell(metaRow1, "Loan ID", loan.Id);
        AddLabelValueCell(metaRow1, "Date", loan.StartDate);
        AddLabelValueCell(metaRow1, "Status", loan.Status);
        metaGroup.Rows.Add(metaRow1);

        var metaRow2 = new TableRow();
        AddLabelValueCell(metaRow2, "Interest Rate", $"{loan.InterestRate}% / month");
        AddLabelValueCell(metaRow2, "Monthly Interest", $"₹{(loan.PrincipalAmount * loan.InterestRate / 100):N2}");
        AddLabelValueCell(metaRow2, "Gov ID", loan.GovId);
        metaGroup.Rows.Add(metaRow2);

        metaTable.RowGroups.Add(metaGroup);
        SetTableBorder(metaTable);
        doc.Blocks.Add(metaTable);

        // ═══ 4. CUSTOMER DETAILS ═══
        var custTable = new Table { CellSpacing = 0 };
        custTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var custGroup = new TableRowGroup();
        var custRow = new TableRow();
        var custCell = new TableCell();
        custCell.BorderBrush = borderBrush;
        custCell.BorderThickness = new Thickness(1, 0, 1, 1);
        custCell.Padding = new Thickness(8, 4, 8, 6);

        custCell.Blocks.Add(new Paragraph(new Run("Details of Pledger"))
        {
            FontSize = 10, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 4),
            TextDecorations = TextDecorations.Underline
        });

        var custInfo = new Paragraph { Margin = new Thickness(0), FontSize = 10, LineHeight = 18 };
        custInfo.Inlines.Add(new Run("Name: ") { Foreground = Brushes.Gray });
        custInfo.Inlines.Add(new Run(loan.CustomerName) { FontWeight = FontWeights.Bold });
        custInfo.Inlines.Add(new Run("\nPhone: ") { Foreground = Brushes.Gray });
        custInfo.Inlines.Add(new Run(loan.CustomerPhone));
        custInfo.Inlines.Add(new Run("\nAddress: ") { Foreground = Brushes.Gray });
        custInfo.Inlines.Add(new Run(loan.CustomerAddress));
        custInfo.Inlines.Add(new Run("\nGov ID: ") { Foreground = Brushes.Gray });
        custInfo.Inlines.Add(new Run(loan.GovId));
        custCell.Blocks.Add(custInfo);

        custRow.Cells.Add(custCell);
        custGroup.Rows.Add(custRow);
        custTable.RowGroups.Add(custGroup);
        doc.Blocks.Add(custTable);

        // ═══ 5. PLEDGED ITEM TABLE ═══
        var itemTable = new Table { CellSpacing = 0 };
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(35) });      // Sr
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) }); // Product
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(70) });      // Metal
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(60) });      // Purity
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(80) });      // Weight
        itemTable.Columns.Add(new TableColumn { Width = new GridLength(120) });     // Principal

        var itemGroup = new TableRowGroup();

        // Header
        var hdr = new TableRow { Background = headerBg };
        AddItemCell(hdr, "Sr", FontWeights.Bold, TextAlignment.Center, 8);
        AddItemCell(hdr, "Product Description", FontWeights.Bold, TextAlignment.Left, 8);
        AddItemCell(hdr, "Metal", FontWeights.Bold, TextAlignment.Center, 8);
        AddItemCell(hdr, "Purity", FontWeights.Bold, TextAlignment.Center, 8);
        AddItemCell(hdr, "Weight (g)", FontWeights.Bold, TextAlignment.Right, 8);
        AddItemCell(hdr, "Principal (₹)", FontWeights.Bold, TextAlignment.Right, 8);
        itemGroup.Rows.Add(hdr);

        // Data
        var dRow = new TableRow { Background = lightBg };
        AddItemCell(dRow, "1", FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(dRow, loan.ProductDescription, FontWeights.Normal, TextAlignment.Left, 9);
        AddItemCell(dRow, loan.MetalType, FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(dRow, loan.Purity, FontWeights.Normal, TextAlignment.Center, 9);
        AddItemCell(dRow, $"{loan.Weight:N3}", FontWeights.Normal, TextAlignment.Right, 9);
        AddItemCell(dRow, $"{loan.PrincipalAmount:N2}", FontWeights.Normal, TextAlignment.Right, 9);
        itemGroup.Rows.Add(dRow);

        // Interest summary row
        var intRow = new TableRow();
        intRow.Cells.Add(MakeSpanCell($"Interest @ {loan.InterestRate}% per month", 5, TextAlignment.Right, FontWeights.Normal, 9));
        AddItemCell(intRow, $"₹{(loan.PrincipalAmount * loan.InterestRate / 100):N2}/month", FontWeights.Bold, TextAlignment.Right, 9);
        itemGroup.Rows.Add(intRow);

        // Principal total row
        var totalRow = new TableRow { Background = darkHeaderBg };
        var totalLabel = new TableCell(new Paragraph(new Run("Total Principal"))
        {
            FontSize = 11, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(4, 3, 4, 3)
        }) { ColumnSpan = 5 };
        totalLabel.BorderBrush = borderBrush;
        totalLabel.BorderThickness = new Thickness(0.5);
        totalRow.Cells.Add(totalLabel);

        var totalVal = new TableCell(new Paragraph(new Run($"₹ {loan.PrincipalAmount:N2}"))
        {
            FontSize = 12, FontWeight = FontWeights.Bold, TextAlignment = TextAlignment.Right,
            Foreground = Brushes.White, Margin = new Thickness(4, 3, 4, 3)
        });
        totalVal.BorderBrush = borderBrush;
        totalVal.BorderThickness = new Thickness(0.5);
        totalRow.Cells.Add(totalVal);
        itemGroup.Rows.Add(totalRow);

        itemTable.RowGroups.Add(itemGroup);
        SetTableBorder(itemTable);
        doc.Blocks.Add(itemTable);

        // ═══ 6. AMOUNT IN WORDS ═══
        var wordsTable = new Table { CellSpacing = 0 };
        wordsTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var wordsGroup = new TableRowGroup();
        var wordsRow = new TableRow { Background = accentBg };
        var wordsCell = new TableCell();
        wordsCell.BorderBrush = borderBrush;
        wordsCell.BorderThickness = new Thickness(1, 0, 1, 1);
        wordsCell.Padding = new Thickness(8, 4, 8, 4);
        var wordsP = new Paragraph { FontSize = 9 };
        wordsP.Inlines.Add(new Run("Principal in words: ") { FontWeight = FontWeights.Bold });
        wordsP.Inlines.Add(new Run(NumberToWords((long)Math.Round(loan.PrincipalAmount)) + " Rupees Only") { FontStyle = FontStyles.Italic, FontWeight = FontWeights.SemiBold });
        wordsCell.Blocks.Add(wordsP);
        wordsRow.Cells.Add(wordsCell);
        wordsGroup.Rows.Add(wordsRow);
        wordsTable.RowGroups.Add(wordsGroup);
        doc.Blocks.Add(wordsTable);

        // ═══ 7. TERMS ═══
        var termsTable = new Table { CellSpacing = 0 };
        termsTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var termsGroup = new TableRowGroup();
        var termsRow = new TableRow();
        var termsCell = new TableCell();
        termsCell.BorderBrush = borderBrush;
        termsCell.BorderThickness = new Thickness(1, 0, 1, 1);
        termsCell.Padding = new Thickness(6, 4, 6, 4);

        termsCell.Blocks.Add(new Paragraph(new Run("Terms & Conditions:"))
        {
            FontSize = 9, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 4)
        });

        var terms = new List { FontSize = 8, Foreground = Brushes.DimGray, MarkerStyle = TextMarkerStyle.Decimal, Padding = new Thickness(16, 0, 0, 0) };
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Article must be redeemed within 12 months from the date of loan."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Interest is charged monthly and must be paid on time."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Pledged article may be auctioned if loan is not repaid within due date."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Customer must bring this receipt at the time of redemption."))));
        terms.ListItems.Add(new ListItem(new Paragraph(new Run("Shop is not responsible for natural wear due to storage."))));
        termsCell.Blocks.Add(terms);
        termsRow.Cells.Add(termsCell);
        termsGroup.Rows.Add(termsRow);
        termsTable.RowGroups.Add(termsGroup);
        doc.Blocks.Add(termsTable);

        // ═══ 8. SIGNATURE AREA ═══
        var sigTable = new Table { CellSpacing = 0 };
        sigTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        sigTable.Columns.Add(new TableColumn { Width = new GridLength(1, GridUnitType.Star) });
        var sigGroup = new TableRowGroup();
        var sigRow = new TableRow();

        var leftSig = new TableCell();
        leftSig.BorderBrush = borderBrush;
        leftSig.BorderThickness = new Thickness(1, 0, 0.5, 1);
        leftSig.Padding = new Thickness(8, 4, 8, 4);
        leftSig.Blocks.Add(new Paragraph(new Run("Pledger's Signature"))
        {
            FontSize = 9, FontWeight = FontWeights.Bold, Margin = new Thickness(0, 0, 0, 30)
        });
        leftSig.Blocks.Add(new Paragraph(new Run("__________________________"))
        {
            FontSize = 8, Foreground = Brushes.Gray
        });
        sigRow.Cells.Add(leftSig);

        var rightSig = new TableCell();
        rightSig.BorderBrush = borderBrush;
        rightSig.BorderThickness = new Thickness(0.5, 0, 1, 1);
        rightSig.Padding = new Thickness(8, 4, 8, 4);
        var rP1 = new Paragraph { TextAlignment = TextAlignment.Right, Margin = new Thickness(0, 0, 0, 4) };
        rP1.Inlines.Add(new Run("For, ") { FontSize = 9, Foreground = Brushes.Gray });
        rP1.Inlines.Add(new Run("ASHISH JEWELLERS") { FontSize = 10, FontWeight = FontWeights.Bold });
        rightSig.Blocks.Add(rP1);
        rightSig.Blocks.Add(new Paragraph(new Run("")) { Margin = new Thickness(0, 0, 0, 16) });
        var rP2 = new Paragraph { TextAlignment = TextAlignment.Right };
        rP2.Inlines.Add(new Run("(Authorised Signatory)") { FontSize = 8, Foreground = Brushes.Gray });
        rightSig.Blocks.Add(rP2);
        sigRow.Cells.Add(rightSig);

        sigGroup.Rows.Add(sigRow);
        sigTable.RowGroups.Add(sigGroup);
        doc.Blocks.Add(sigTable);

        // ═══ 9. FOOTER ═══
        var footerText = isDraft ? "⚠ DRAFT — This is NOT an official loan receipt." : "This is a computer generated loan receipt.";
        doc.Blocks.Add(new Paragraph(new Run(footerText))
        {
            FontSize = 7, Foreground = Brushes.Gray, TextAlignment = TextAlignment.Center,
            Margin = new Thickness(0, 6, 0, 0)
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
