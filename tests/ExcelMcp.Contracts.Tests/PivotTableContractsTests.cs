using ExcelMcp.Contracts;
using Xunit;

namespace ExcelMcp.Contracts.Tests;

public sealed class PivotTableContractsTests
{
    [Fact]
    public void PivotTableArguments_DefaultValues()
    {
        var args = new PivotTableArguments("Sheet1");

        Assert.Equal("Sheet1", args.Worksheet);
        Assert.Null(args.PivotTable);
        Assert.True(args.IncludeFilters);
        Assert.Equal(100, args.MaxRows);
    }

    [Fact]
    public void PivotTableArguments_AllParameters()
    {
        var args = new PivotTableArguments(
            Worksheet: "Analysis",
            PivotTable: "SalesPivot",
            IncludeFilters: false,
            MaxRows: 50);

        Assert.Equal("Analysis", args.Worksheet);
        Assert.Equal("SalesPivot", args.PivotTable);
        Assert.False(args.IncludeFilters);
        Assert.Equal(50, args.MaxRows);
    }

    [Fact]
    public void PivotTableResult_EmptyResult()
    {
        var result = new PivotTableResult(Array.Empty<PivotTableInfo>());

        Assert.Empty(result.PivotTables);
    }

    [Fact]
    public void PivotTableResult_WithPivotTables()
    {
        var pivotInfo = new PivotTableInfo(
            Name: "SalesPivot",
            WorksheetName: "Analysis",
            SourceWorksheet: "RawData",
            SourceRange: "A1:F100",
            RowFields: Array.Empty<PivotFieldInfo>(),
            ColumnFields: Array.Empty<PivotFieldInfo>(),
            DataFields: Array.Empty<PivotFieldInfo>(),
            FilterFields: Array.Empty<PivotFieldInfo>(),
            Data: Array.Empty<PivotDataRow>());

        var result = new PivotTableResult(new[] { pivotInfo });

        Assert.Single(result.PivotTables);
        Assert.Equal("SalesPivot", result.PivotTables[0].Name);
    }

    [Fact]
    public void PivotTableInfo_AllFields()
    {
        var rowFields = new[] { new PivotFieldInfo("Region", "RegionSource", "Row") };
        var columnFields = new[] { new PivotFieldInfo("Quarter", "QuarterSource", "Column") };
        var dataFields = new[] { new PivotFieldInfo("SumOfSales", "Sales", "Sum") };
        var filterFields = new[] { new PivotFieldInfo("Year", "YearSource", "Filter") };
        var data = new[] { new PivotDataRow(new Dictionary<string, string?> { { "Value", "100" } }) };

        var info = new PivotTableInfo(
            Name: "SalesPivot",
            WorksheetName: "Analysis",
            SourceWorksheet: "RawData",
            SourceRange: "A1:F100",
            RowFields: rowFields,
            ColumnFields: columnFields,
            DataFields: dataFields,
            FilterFields: filterFields,
            Data: data);

        Assert.Equal("SalesPivot", info.Name);
        Assert.Equal("Analysis", info.WorksheetName);
        Assert.Equal("RawData", info.SourceWorksheet);
        Assert.Equal("A1:F100", info.SourceRange);
        Assert.Single(info.RowFields);
        Assert.Single(info.ColumnFields);
        Assert.Single(info.DataFields);
        Assert.Single(info.FilterFields);
        Assert.Single(info.Data);
    }

    [Fact]
    public void PivotFieldInfo_CanBeCreated()
    {
        var field = new PivotFieldInfo("TotalSales", "Sales", "Sum");

        Assert.Equal("TotalSales", field.Name);
        Assert.Equal("Sales", field.SourceName);
        Assert.Equal("Sum", field.Function);
    }

    [Fact]
    public void PivotDataRow_CanBeCreated()
    {
        var values = new Dictionary<string, string?>
        {
            { "Region", "North" },
            { "Sales", "5000" },
            { "Empty", null }
        };

        var row = new PivotDataRow(values);

        Assert.Equal(3, row.Values.Count);
        Assert.Equal("North", row.Values["Region"]);
        Assert.Equal("5000", row.Values["Sales"]);
        Assert.Null(row.Values["Empty"]);
    }

    [Fact]
    public void PivotTableArguments_RecordEquality()
    {
        var args1 = new PivotTableArguments("Sheet1", "Pivot1", true, 100);
        var args2 = new PivotTableArguments("Sheet1", "Pivot1", true, 100);

        Assert.Equal(args1, args2);
    }

    [Fact]
    public void PivotFieldInfo_RecordEquality()
    {
        var field1 = new PivotFieldInfo("Name", "Source", "Function");
        var field2 = new PivotFieldInfo("Name", "Source", "Function");

        Assert.Equal(field1, field2);
    }
}
