using ExcelMcp.Contracts;
using Xunit;

namespace ExcelMcp.Contracts.Tests;

public sealed class ExcelSearchContractsTests
{
    [Fact]
    public void ExcelSearchArguments_DefaultValues()
    {
        var args = new ExcelSearchArguments("test query");

        Assert.Equal("test query", args.Query);
        Assert.Null(args.Worksheet);
        Assert.Null(args.Table);
        Assert.Null(args.Limit);
        Assert.False(args.CaseSensitive);
    }

    [Fact]
    public void ExcelSearchArguments_AllParameters()
    {
        var args = new ExcelSearchArguments(
            Query: "search term",
            Worksheet: "Sheet1",
            Table: "Table1",
            Limit: 50,
            CaseSensitive: true);

        Assert.Equal("search term", args.Query);
        Assert.Equal("Sheet1", args.Worksheet);
        Assert.Equal("Table1", args.Table);
        Assert.Equal(50, args.Limit);
        Assert.True(args.CaseSensitive);
    }

    [Fact]
    public void ExcelSearchResult_EmptyResult()
    {
        var result = new ExcelSearchResult(Array.Empty<ExcelRowResult>(), false);

        Assert.Empty(result.Rows);
        Assert.False(result.HasMore);
    }

    [Fact]
    public void ExcelSearchResult_WithRows()
    {
        var values = new Dictionary<string, string?> { { "Name", "John" }, { "Age", "30" } };
        var rows = new List<ExcelRowResult>
        {
            new("Sheet1", "Table1", 2, values)
        };

        var result = new ExcelSearchResult(rows, true);

        Assert.Single(result.Rows);
        Assert.True(result.HasMore);
    }

    [Fact]
    public void ExcelRowResult_CanBeCreated()
    {
        var values = new Dictionary<string, string?>
        {
            { "Column1", "Value1" },
            { "Column2", null }
        };

        var row = new ExcelRowResult("Sheet1", "Table1", 5, values);

        Assert.Equal("Sheet1", row.WorksheetName);
        Assert.Equal("Table1", row.TableName);
        Assert.Equal(5, row.RowNumber);
        Assert.Equal(2, row.Values.Count);
        Assert.Equal("Value1", row.Values["Column1"]);
        Assert.Null(row.Values["Column2"]);
    }

    [Fact]
    public void ExcelRowResult_WithoutTable()
    {
        var values = new Dictionary<string, string?> { { "Header", "Data" } };
        var row = new ExcelRowResult("Sheet1", null, 10, values);

        Assert.Equal("Sheet1", row.WorksheetName);
        Assert.Null(row.TableName);
        Assert.Equal(10, row.RowNumber);
    }

    [Fact]
    public void ExcelSearchArguments_RecordEquality()
    {
        var args1 = new ExcelSearchArguments("query", "Sheet1", "Table1", 10, true);
        var args2 = new ExcelSearchArguments("query", "Sheet1", "Table1", 10, true);

        Assert.Equal(args1, args2);
    }

    [Fact]
    public void ExcelSearchArguments_RecordInequality()
    {
        var args1 = new ExcelSearchArguments("query1");
        var args2 = new ExcelSearchArguments("query2");

        Assert.NotEqual(args1, args2);
    }
}
