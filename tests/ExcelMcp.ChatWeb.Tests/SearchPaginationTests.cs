using System.Linq;
using ClosedXML.Excel;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.ChatWeb.Tests;

public sealed class SearchPaginationTests
{
    [Fact]
    public async Task SearchAsync_ReturnsPagedResultsAndCursor()
    {
        var workbookPath = CreateWorkbook(5);
        try
        {
            var service = new ExcelWorkbookService(workbookPath);
            var firstPageArgs = new ExcelSearchArguments("match", Worksheet: "Sheet1", Limit: 2);

            var firstPage = await service.SearchAsync(firstPageArgs, CancellationToken.None);

            Assert.Equal(2, firstPage.Rows.Count);
            Assert.True(firstPage.HasMore);
            Assert.Equal("2", firstPage.NextCursor);
            Assert.All(firstPage.Rows, row => Assert.Equal("Sheet1", row.WorksheetName));

            var secondPageArgs = firstPageArgs with { Cursor = firstPage.NextCursor };
            var secondPage = await service.SearchAsync(secondPageArgs, CancellationToken.None);

            Assert.Equal(2, secondPage.Rows.Count);
            Assert.True(secondPage.HasMore);
            Assert.Equal("4", secondPage.NextCursor);
            Assert.Equal("match-3", secondPage.Rows.First().Values["Header"]);

            var finalPageArgs = firstPageArgs with { Cursor = secondPage.NextCursor };
            var finalPage = await service.SearchAsync(finalPageArgs, CancellationToken.None);

            Assert.Single(finalPage.Rows);
            Assert.False(finalPage.HasMore);
            Assert.Null(finalPage.NextCursor);
            Assert.Equal("match-5", finalPage.Rows.Single().Values["Header"]);
        }
        finally
        {
            File.Delete(workbookPath);
        }
    }

    [Fact]
    public async Task SearchAsync_ReturnsEmptyPageWhenCursorBeyondResults()
    {
        var workbookPath = CreateWorkbook(3);
        try
        {
            var service = new ExcelWorkbookService(workbookPath);
            var args = new ExcelSearchArguments("match", Worksheet: "Sheet1", Limit: 2, Cursor: "10");

            var result = await service.SearchAsync(args, CancellationToken.None);

            Assert.Empty(result.Rows);
            Assert.False(result.HasMore);
            Assert.Null(result.NextCursor);
        }
        finally
        {
            File.Delete(workbookPath);
        }
    }

    private static string CreateWorkbook(int rowCount)
    {
        var path = Path.Combine(Path.GetTempPath(), $"excelmcp-search-{Guid.NewGuid():N}.xlsx");
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell(1, 1).Value = "Header";

        for (var i = 0; i < rowCount; i++)
        {
            sheet.Cell(i + 2, 1).Value = $"match-{i + 1}";
        }

        workbook.SaveAs(path);
        return path;
    }
}
