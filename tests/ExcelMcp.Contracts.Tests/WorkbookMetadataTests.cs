using ExcelMcp.Contracts;
using Xunit;

namespace ExcelMcp.Contracts.Tests;

public sealed class WorkbookMetadataTests
{
    [Fact]
    public void WorkbookMetadata_CanBeCreated()
    {
        var worksheets = new List<WorksheetMetadata>
        {
            new("Sheet1", Array.Empty<TableMetadata>(), new[] { "Column1", "Column2" }, Array.Empty<PivotTableMetadata>())
        };

        var metadata = new WorkbookMetadata("/path/to/workbook.xlsx", worksheets, DateTimeOffset.UtcNow);

        Assert.Equal("/path/to/workbook.xlsx", metadata.WorkbookPath);
        Assert.Single(metadata.Worksheets);
        Assert.NotEqual(default, metadata.LastLoadedUtc);
    }

    [Fact]
    public void WorksheetMetadata_CanBeCreated()
    {
        var tables = new List<TableMetadata>
        {
            new("Table1", "Sheet1", new[] { "Col1", "Col2" }, 100)
        };
        var headers = new[] { "Header1", "Header2" };
        var pivotTables = new List<PivotTableMetadata>
        {
            new("Pivot1", "Sheet1", "A1:D10", 2, 1, 1)
        };

        var worksheet = new WorksheetMetadata("Sheet1", tables, headers, pivotTables);

        Assert.Equal("Sheet1", worksheet.Name);
        Assert.Single(worksheet.Tables);
        Assert.Equal(2, worksheet.ColumnHeaders.Count);
        Assert.Single(worksheet.PivotTables);
    }

    [Fact]
    public void TableMetadata_CanBeCreated()
    {
        var headers = new[] { "Name", "Age", "Department" };
        var table = new TableMetadata("EmployeeTable", "Sheet1", headers, 50);

        Assert.Equal("EmployeeTable", table.Name);
        Assert.Equal("Sheet1", table.WorksheetName);
        Assert.Equal(3, table.ColumnHeaders.Count);
        Assert.Equal(50, table.RowCount);
    }

    [Fact]
    public void PivotTableMetadata_CanBeCreated()
    {
        var pivot = new PivotTableMetadata("SalesPivot", "Analysis", "SalesData!A1:F100", 3, 2, 1);

        Assert.Equal("SalesPivot", pivot.Name);
        Assert.Equal("Analysis", pivot.WorksheetName);
        Assert.Equal("SalesData!A1:F100", pivot.SourceRange);
        Assert.Equal(3, pivot.RowFieldCount);
        Assert.Equal(2, pivot.ColumnFieldCount);
        Assert.Equal(1, pivot.DataFieldCount);
    }

    [Fact]
    public void WorkbookMetadata_RecordEquality()
    {
        var worksheets = Array.Empty<WorksheetMetadata>().ToList();
        var time = DateTimeOffset.UtcNow;

        var metadata1 = new WorkbookMetadata("/path.xlsx", worksheets, time);
        var metadata2 = new WorkbookMetadata("/path.xlsx", worksheets, time);

        Assert.Equal(metadata1, metadata2);
    }

    [Fact]
    public void WorksheetMetadata_EmptyCollections()
    {
        var worksheet = new WorksheetMetadata(
            "EmptySheet",
            Array.Empty<TableMetadata>(),
            Array.Empty<string>(),
            Array.Empty<PivotTableMetadata>());

        Assert.Equal("EmptySheet", worksheet.Name);
        Assert.Empty(worksheet.Tables);
        Assert.Empty(worksheet.ColumnHeaders);
        Assert.Empty(worksheet.PivotTables);
    }
}
