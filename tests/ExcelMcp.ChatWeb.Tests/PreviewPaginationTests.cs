using System.Linq;
using ClosedXML.Excel;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.ChatWeb.Tests;

public sealed class PreviewPaginationTests
{
    [Fact]
    public async Task PreviewAsync_PaginatesWorksheetRows()
    {
        var workbookPath = CreateWorksheetWorkbook(5);
        try
        {
            var service = new ExcelWorkbookService(workbookPath);
            var firstPageArgs = new ExcelPreviewArguments("Sheet1", Rows: 2);

            var firstPage = await service.PreviewAsync(firstPageArgs, CancellationToken.None);

            Assert.Equal("Sheet1", firstPage.Worksheet);
            Assert.Null(firstPage.Table);
            Assert.Equal(2, firstPage.Rows.Count);
            Assert.Equal(0, firstPage.Offset);
            Assert.True(firstPage.HasMore);
            Assert.Equal("2", firstPage.NextCursor);
            Assert.Equal(2, firstPage.Rows.First().RowNumber);
            Assert.Contains("Header", firstPage.Csv);

            var secondPageArgs = firstPageArgs with { Cursor = firstPage.NextCursor };
            var secondPage = await service.PreviewAsync(secondPageArgs, CancellationToken.None);

            Assert.Equal(2, secondPage.Rows.Count);
            Assert.Equal(2, secondPage.Offset);
            Assert.True(secondPage.HasMore);
            Assert.Equal("4", secondPage.NextCursor);
            Assert.Equal("value-4", secondPage.Rows.Last().Values[0]);

            var finalPageArgs = firstPageArgs with { Cursor = secondPage.NextCursor };
            var finalPage = await service.PreviewAsync(finalPageArgs, CancellationToken.None);

            Assert.Single(finalPage.Rows);
            Assert.Equal(4, finalPage.Offset);
            Assert.False(finalPage.HasMore);
            Assert.Null(finalPage.NextCursor);
            Assert.Equal("value-5", finalPage.Rows.Single().Values[0]);
        }
        finally
        {
            File.Delete(workbookPath);
        }
    }

    [Fact]
    public async Task PreviewAsync_PaginatesTableRows()
    {
        var workbookPath = CreateTableWorkbook(4);
        try
        {
            var service = new ExcelWorkbookService(workbookPath);
            var args = new ExcelPreviewArguments("Sheet1", Table: "SampleTable", Rows: 2);

            var firstPage = await service.PreviewAsync(args, CancellationToken.None);

            Assert.Equal("Sheet1", firstPage.Worksheet);
            Assert.Equal("SampleTable", firstPage.Table);
            Assert.Equal(new[] { "Id", "Name" }, firstPage.Headers);
            Assert.Equal(2, firstPage.Rows.Count);
            Assert.True(firstPage.HasMore);
            Assert.Equal("2", firstPage.NextCursor);
            Assert.Equal("Item-1", firstPage.Rows.First().Values[1]);

            var secondPage = await service.PreviewAsync(args with { Cursor = firstPage.NextCursor }, CancellationToken.None);

            Assert.Equal(2, secondPage.Rows.Count);
            Assert.False(string.IsNullOrWhiteSpace(secondPage.Csv));
            Assert.False(secondPage.HasMore);
            Assert.Null(secondPage.NextCursor);
            Assert.Equal("Item-4", secondPage.Rows.Last().Values[1]);
        }
        finally
        {
            File.Delete(workbookPath);
        }
    }

    private static string CreateWorksheetWorkbook(int rowCount)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-preview-sheet-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell(1, 1).Value = "Header";

        for (var i = 0; i < rowCount; i++)
        {
            sheet.Cell(i + 2, 1).Value = $"value-{i + 1}";
        }

        workbook.SaveAs(path);
        return path;
    }

    private static string CreateTableWorkbook(int rowCount)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-preview-table-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell(1, 1).Value = "Id";
        sheet.Cell(1, 2).Value = "Name";

        for (var i = 0; i < rowCount; i++)
        {
            sheet.Cell(i + 2, 1).Value = i + 1;
            sheet.Cell(i + 2, 2).Value = $"Item-{i + 1}";
        }

        var tableRange = sheet.Range(1, 1, rowCount + 1, 2);
        tableRange.CreateTable("SampleTable");

        workbook.SaveAs(path);
        return path;
    }
}
