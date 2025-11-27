using ExcelMcp.Contracts;
using Xunit;

namespace ExcelMcp.Contracts.Tests;

public sealed class ResourceContractsTests
{
    [Fact]
    public void ExcelResourceDescriptor_CanBeCreated()
    {
        var uri = new Uri("excel://workbook");
        var descriptor = new ExcelResourceDescriptor(uri, "Workbook", "Main workbook", "application/json");

        Assert.Equal(uri, descriptor.Uri);
        Assert.Equal("Workbook", descriptor.Name);
        Assert.Equal("Main workbook", descriptor.Description);
        Assert.Equal("application/json", descriptor.MimeType);
    }

    [Fact]
    public void ExcelResourceDescriptor_NullDescription()
    {
        var uri = new Uri("excel://worksheet/Sheet1");
        var descriptor = new ExcelResourceDescriptor(uri, "Sheet1", null, "text/csv");

        Assert.Equal(uri, descriptor.Uri);
        Assert.Equal("Sheet1", descriptor.Name);
        Assert.Null(descriptor.Description);
        Assert.Equal("text/csv", descriptor.MimeType);
    }

    [Fact]
    public void ExcelResourceContent_CanBeCreated()
    {
        var uri = new Uri("excel://workbook");
        var content = new ExcelResourceContent(uri, "application/json", "{\"key\":\"value\"}");

        Assert.Equal(uri, content.Uri);
        Assert.Equal("application/json", content.MimeType);
        Assert.Equal("{\"key\":\"value\"}", content.Text);
    }

    [Fact]
    public void ExcelResourceContent_CsvContent()
    {
        var uri = new Uri("excel://worksheet/Sheet1");
        var csvData = "Name,Age\nJohn,30\nJane,25";
        var content = new ExcelResourceContent(uri, "text/csv", csvData);

        Assert.Equal(uri, content.Uri);
        Assert.Equal("text/csv", content.MimeType);
        Assert.Contains("Name,Age", content.Text);
    }

    [Fact]
    public void ExcelResourceDescriptor_RecordEquality()
    {
        var uri = new Uri("excel://workbook");
        var desc1 = new ExcelResourceDescriptor(uri, "Name", "Desc", "mime");
        var desc2 = new ExcelResourceDescriptor(uri, "Name", "Desc", "mime");

        Assert.Equal(desc1, desc2);
    }

    [Fact]
    public void ExcelResourceContent_RecordEquality()
    {
        var uri = new Uri("excel://workbook");
        var content1 = new ExcelResourceContent(uri, "text/csv", "data");
        var content2 = new ExcelResourceContent(uri, "text/csv", "data");

        Assert.Equal(content1, content2);
    }

    [Fact]
    public void ExcelResourceDescriptor_RecordInequality()
    {
        var uri1 = new Uri("excel://workbook");
        var uri2 = new Uri("excel://worksheet/Sheet1");

        var desc1 = new ExcelResourceDescriptor(uri1, "Name1", null, "mime");
        var desc2 = new ExcelResourceDescriptor(uri2, "Name2", null, "mime");

        Assert.NotEqual(desc1, desc2);
    }
}
