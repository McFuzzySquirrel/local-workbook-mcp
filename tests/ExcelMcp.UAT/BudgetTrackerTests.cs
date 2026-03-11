using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.UAT;

/// <summary>
/// User Acceptance Tests — BudgetTracker.xlsx
///
/// Covers financial tracking scenarios.
/// Re-generate test data with: scripts/create-sample-workbooks.ps1
/// </summary>
public sealed class BudgetTrackerTests
{
    private static readonly string WorkbookPath = TestData.GetPath("BudgetTracker.xlsx");

    // ── Structure ────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-BT-01: Workbook has two worksheets: Income and Expenses")]
    public async Task WorkbookHasTwoExpectedSheets()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheetNames = metadata.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Contains("Income", sheetNames);
        Assert.Contains("Expenses", sheetNames);
        Assert.Equal(2, sheetNames.Count);
    }

    [Fact(DisplayName = "UA-BT-02: IncomeTable has six rows")]
    public async Task IncomeTableHasSixRows()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheet = metadata.Worksheets.Single(ws => ws.Name == "Income");
        var table = sheet.Tables.Single(t => t.Name == "IncomeTable");

        Assert.Equal(6, table.RowCount);
        Assert.Contains("Source", table.ColumnHeaders);
        Assert.Contains("Category", table.ColumnHeaders);
        Assert.Contains("Amount", table.ColumnHeaders);
    }

    [Fact(DisplayName = "UA-BT-03: ExpensesTable has eight rows")]
    public async Task ExpensesTableHasEightRows()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheet = metadata.Worksheets.Single(ws => ws.Name == "Expenses");
        var table = sheet.Tables.Single(t => t.Name == "ExpensesTable");

        Assert.Equal(8, table.RowCount);
    }

    // ── Search: Income ───────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-BT-04: Client A appears exactly twice in the Income table")]
    public async Task ClientA_AppearsTwiceInIncome()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Client A", Worksheet: "Income", Table: "IncomeTable"),
            CancellationToken.None);

        Assert.Equal(2, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Client A", row.Values["Source"]));
    }

    [Fact(DisplayName = "UA-BT-05: Client B income totals 15500 (8000 + 7500)")]
    public async Task ClientB_IncomeTotals15500()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Client B", Worksheet: "Income", Table: "IncomeTable"),
            CancellationToken.None);

        Assert.Equal(2, result.Rows.Count);
        var total = result.Rows.Sum(row => double.Parse(row.Values["Amount"]!));
        Assert.Equal(15500.0, total);
    }

    [Fact(DisplayName = "UA-BT-06: Total income across all six entries is 31500")]
    public async Task TotalIncome_Is31500()
    {
        var service = new ExcelWorkbookService(WorkbookPath);

        // Each client appears at least once — use a broad search across all rows
        // by previewing the whole sheet and parsing, or search each client.
        // We use a workaround: preview the CSV content and parse.
        var uri = ExcelResourceUri.CreateTableUri("Income", "IncomeTable");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 20);

        // Parse the CSV: header row + 6 data rows
        var lines = content.Text!
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Skip(1) // skip header
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToList();

        Assert.Equal(6, lines.Count);

        var total = lines.Sum(line =>
        {
            var cols = line.Split(',');
            // Amount is the last (4th) column
            return double.TryParse(cols[^1].Trim(), out var v) ? v : 0;
        });

        Assert.Equal(31500.0, total);
    }

    [Fact(DisplayName = "UA-BT-07: Consulting category appears in income — Client A twice")]
    public async Task ConsultingCategory_AppearsInIncome()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Consulting", Worksheet: "Income", Table: "IncomeTable"),
            CancellationToken.None);

        Assert.Equal(2, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Consulting", row.Values["Category"]));
    }

    // ── Search: Expenses ─────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-BT-08: Cloud Provider appears twice in expenses (both Software)")]
    public async Task CloudProvider_AppearsTwiceInExpenses()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Cloud Provider", Worksheet: "Expenses", Table: "ExpensesTable"),
            CancellationToken.None);

        Assert.Equal(2, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Software", row.Values["Category"]));
    }

    [Fact(DisplayName = "UA-BT-09: Largest single expense is Office Rent at 3000")]
    public async Task OfficRent_IsLargestExpense()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Rent", Worksheet: "Expenses", Table: "ExpensesTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        Assert.Equal("3000", result.Rows[0].Values["Amount"]);
        Assert.Equal("Rent", result.Rows[0].Values["Category"]);
    }

    [Fact(DisplayName = "UA-BT-10: Total expenses across all eight rows is 10535")]
    public async Task TotalExpenses_Is10535()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var uri = ExcelResourceUri.CreateTableUri("Expenses", "ExpensesTable");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 20);

        var lines = content.Text!
            .Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Skip(1)
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToList();

        Assert.Equal(8, lines.Count);

        var total = lines.Sum(line =>
        {
            var cols = line.Split(',');
            return double.TryParse(cols[^1].Trim(), out var v) ? v : 0;
        });

        Assert.Equal(10535.0, total);
    }

    // ── Search: Cross-sheet ──────────────────────────────────────────────────

    [Fact(DisplayName = "UA-BT-11: Searching 'Software' across whole workbook finds expense rows only")]
    public async Task SearchSoftware_FinadsExpenseRows()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        // Search without constraining to a table — hits both sheets
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Software"),
            CancellationToken.None);

        Assert.True(result.Rows.Count >= 2);
        Assert.All(result.Rows, row =>
            Assert.Equal("Expenses", row.WorksheetName));
    }

    // ── Preview ──────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-BT-12: Previewing Expenses sheet returns CSV with Vendor and Category columns")]
    public async Task PreviewExpenses_HasVendorAndCategoryColumns()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var uri = ExcelResourceUri.CreateWorksheetUri("Expenses");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 5);

        Assert.Equal("text/csv", content.MimeType);
        Assert.NotNull(content.Text);
        Assert.Contains("Vendor", content.Text);
        Assert.Contains("Category", content.Text);
    }
}
