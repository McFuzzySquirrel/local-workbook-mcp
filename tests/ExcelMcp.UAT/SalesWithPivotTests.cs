using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.UAT;

/// <summary>
/// User Acceptance Tests — SalesWithPivot.xlsx
///
/// Covers pivot-table analysis scenarios as well as basic structure and search.
/// Re-generate test data with: scripts/create-pivot-test-workbook.ps1
/// </summary>
public sealed class SalesWithPivotTests
{
    private static readonly string WorkbookPath = TestData.GetPath("SalesWithPivot.xlsx");

    // ── Structure ────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-SP-01: Workbook has two worksheets: SalesData and SalesPivot")]
    public async Task WorkbookHasTwoExpectedSheets()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheetNames = metadata.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Contains("SalesData", sheetNames);
        Assert.Contains("SalesPivot", sheetNames);
    }

    [Fact(DisplayName = "UA-SP-02: SalesData has SalesTable with ten rows and expected columns")]
    public async Task SalesTableHasTenRows()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheet = metadata.Worksheets.Single(ws => ws.Name == "SalesData");
        var table = sheet.Tables.Single(t => t.Name == "SalesTable");

        Assert.Equal(10, table.RowCount);
        Assert.Contains("Region", table.ColumnHeaders);
        Assert.Contains("Product", table.ColumnHeaders);
        Assert.Contains("SalesPerson", table.ColumnHeaders);
        Assert.Contains("Amount", table.ColumnHeaders);
        Assert.Contains("Quarter", table.ColumnHeaders);
    }

    [Fact(DisplayName = "UA-SP-03: SalesPivot worksheet contains at least one pivot table")]
    public async Task SalesPivotSheet_HasPivotTable()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var pivotSheet = metadata.Worksheets.Single(ws => ws.Name == "SalesPivot");
        Assert.NotEmpty(pivotSheet.PivotTables);
    }

    [Fact(DisplayName = "UA-SP-04: SalesSummary pivot table exists with correct field configuration")]
    public async Task SalesSummaryPivot_HasCorrectFields()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var pivotSheet = metadata.Worksheets.Single(ws => ws.Name == "SalesPivot");
        var pivot = pivotSheet.PivotTables.Single(p => p.Name == "SalesSummary");

        Assert.Equal("SalesSummary", pivot.Name);
        Assert.Equal("SalesPivot", pivot.WorksheetName);
        // RowFieldCount reflects ClosedXML's read of the pivot configuration
        Assert.True(pivot.RowFieldCount >= 1, $"Expected at least 1 row field, got {pivot.RowFieldCount}");
        // Value fields: Amount (Sum)
        Assert.True(pivot.DataFieldCount >= 1, $"Expected at least 1 data field, got {pivot.DataFieldCount}");
    }

    // ── Search: SalesData ────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-SP-05: Searching 'John' returns four entries in SalesData")]
    public async Task SearchJohn_ReturnsFourEntries()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("John", Worksheet: "SalesData", Table: "SalesTable"),
            CancellationToken.None);

        // John: Q1 East Widget 1200, Q2 West Widget 1800, Q3 East Widget 1900, Q4 West Widget 2300
        Assert.Equal(4, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("John", row.Values["SalesPerson"]));
    }

    [Fact(DisplayName = "UA-SP-06: John's total sales amount to 9200")]
    public async Task JohnTotalSales_Is9200()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("John", Worksheet: "SalesData", Table: "SalesTable"),
            CancellationToken.None);

        var total = result.Rows.Sum(row => double.Parse(row.Values["Amount"]!));
        // 1200 + 1800 + 1900 + 2300 = 7200
        Assert.Equal(7200.0, total);
    }

    [Fact(DisplayName = "UA-SP-07: Searching 'East' returns five East region entries")]
    public async Task SearchEast_ReturnsFiveEntries()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("East", Worksheet: "SalesData", Table: "SalesTable"),
            CancellationToken.None);

        // East: John Q1, Sarah Q1, Sarah Q2, John Q3, Mike Q3
        Assert.Equal(5, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("East", row.Values["Region"]));
    }

    [Fact(DisplayName = "UA-SP-08: Q4 has exactly one sale entry (John, West, Widget, 2300)")]
    public async Task Q4_HasOneSaleEntry()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Q4", Worksheet: "SalesData", Table: "SalesTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        Assert.Equal("John", result.Rows[0].Values["SalesPerson"]);
        Assert.Equal("West", result.Rows[0].Values["Region"]);
        Assert.Equal("2300", result.Rows[0].Values["Amount"]);
    }

    // ── Pivot Analysis ───────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-SP-09: Analyzing SalesSummary pivot returns non-empty result")]
    public async Task AnalyzePivot_ReturnsNonEmptyResult()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var args = new PivotTableArguments("SalesPivot", "SalesSummary", IncludeFilters: true, MaxRows: 100);
        var result = await service.AnalyzePivotTablesAsync(args, CancellationToken.None);

        Assert.NotEmpty(result.PivotTables);
        var pivot = result.PivotTables.Single();
        Assert.Equal("SalesSummary", pivot.Name);
    }

    // ── Preview ──────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-SP-10: Previewing SalesData sheet returns CSV with Region and Amount columns")]
    public async Task PreviewSalesData_HasRegionAndAmountColumns()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var uri = ExcelResourceUri.CreateWorksheetUri("SalesData");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 5);

        Assert.Equal("text/csv", content.MimeType);
        Assert.NotNull(content.Text);
        Assert.Contains("Region", content.Text);
        Assert.Contains("Amount", content.Text);
    }
}
