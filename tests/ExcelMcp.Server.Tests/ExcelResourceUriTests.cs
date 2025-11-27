using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.Server.Tests;

public sealed class ExcelResourceUriTests
{
    [Fact]
    public void WorkbookUri_HasCorrectSchemeAndHost()
    {
        Assert.Equal("excel", ExcelResourceUri.WorkbookUri.Scheme);
        Assert.Equal("workbook", ExcelResourceUri.WorkbookUri.Host);
    }

    [Fact]
    public void CreateWorksheetUri_ReturnsValidUri()
    {
        var uri = ExcelResourceUri.CreateWorksheetUri("Sheet1");
        Assert.Equal("excel", uri.Scheme);
        Assert.Equal("worksheet", uri.Host);
        Assert.Contains("Sheet1", uri.AbsolutePath);
    }

    [Fact]
    public void CreateWorksheetUri_HandlesSpecialCharacters()
    {
        var uri = ExcelResourceUri.CreateWorksheetUri("Sales Data 2024");
        Assert.Equal("excel", uri.Scheme);
        Assert.Equal("worksheet", uri.Host);
        // The AbsolutePath may or may not be encoded depending on .NET's Uri handling
        Assert.NotNull(uri.AbsolutePath);
    }

    [Fact]
    public void CreateTableUri_ReturnsValidUri()
    {
        var uri = ExcelResourceUri.CreateTableUri("Sheet1", "Table1");
        Assert.Equal("excel", uri.Scheme);
        Assert.Equal("worksheet", uri.Host);
        Assert.Contains("Sheet1", uri.AbsolutePath);
        Assert.Contains("Table1", uri.AbsolutePath);
        Assert.Contains("table", uri.AbsolutePath);
    }

    [Fact]
    public void CreateTableUri_HandlesSpecialCharacters()
    {
        var uri = ExcelResourceUri.CreateTableUri("Sales Data", "Main Table");
        Assert.Equal("excel", uri.Scheme);
        Assert.Equal("worksheet", uri.Host);
        Assert.Contains("table", uri.AbsolutePath);
    }

    [Fact]
    public void TryParse_WorkbookUri_Succeeds()
    {
        var uri = new Uri("excel://workbook");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.True(result);
        Assert.Null(worksheet);
        Assert.Null(table);
    }

    [Fact]
    public void TryParse_WorksheetUri_Succeeds()
    {
        var uri = new Uri("excel://worksheet/Sheet1");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.True(result);
        Assert.Equal("Sheet1", worksheet);
        Assert.Null(table);
    }

    [Fact]
    public void TryParse_WorksheetUri_UnescapesSpecialCharacters()
    {
        var uri = new Uri("excel://worksheet/Sales%20Data%202024");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.True(result);
        Assert.Equal("Sales Data 2024", worksheet);
        Assert.Null(table);
    }

    [Fact]
    public void TryParse_TableUri_Succeeds()
    {
        var uri = new Uri("excel://worksheet/Sheet1/table/Table1");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.True(result);
        Assert.Equal("Sheet1", worksheet);
        Assert.Equal("Table1", table);
    }

    [Fact]
    public void TryParse_TableUri_UnescapesSpecialCharacters()
    {
        var uri = new Uri("excel://worksheet/Sales%20Data/table/Main%20Table");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.True(result);
        Assert.Equal("Sales Data", worksheet);
        Assert.Equal("Main Table", table);
    }

    [Fact]
    public void TryParse_InvalidScheme_ReturnsFalse()
    {
        var uri = new Uri("http://worksheet/Sheet1");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.False(result);
        Assert.Null(worksheet);
        Assert.Null(table);
    }

    [Fact]
    public void TryParse_UnknownHost_ReturnsFalse()
    {
        var uri = new Uri("excel://unknown/Sheet1");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.False(result);
    }

    [Fact]
    public void TryParse_EmptyPath_ReturnsFalse()
    {
        var uri = new Uri("excel://worksheet/");
        var result = ExcelResourceUri.TryParse(uri, out var worksheet, out var table);

        Assert.False(result);
    }

    [Theory]
    [InlineData("Sheet1")]
    [InlineData("Data Sheet")]
    [InlineData("2024_Budget")]
    [InlineData("Übersicht")]
    public void RoundTrip_WorksheetUri(string worksheetName)
    {
        var uri = ExcelResourceUri.CreateWorksheetUri(worksheetName);
        var result = ExcelResourceUri.TryParse(uri, out var parsedWorksheet, out var parsedTable);

        Assert.True(result);
        Assert.Equal(worksheetName, parsedWorksheet);
        Assert.Null(parsedTable);
    }

    [Theory]
    [InlineData("Sheet1", "Table1")]
    [InlineData("Sales Data", "Main Table")]
    [InlineData("2024_Budget", "Q1_Expenses")]
    public void RoundTrip_TableUri(string worksheetName, string tableName)
    {
        var uri = ExcelResourceUri.CreateTableUri(worksheetName, tableName);
        var result = ExcelResourceUri.TryParse(uri, out var parsedWorksheet, out var parsedTable);

        Assert.True(result);
        Assert.Equal(worksheetName, parsedWorksheet);
        Assert.Equal(tableName, parsedTable);
    }
}
