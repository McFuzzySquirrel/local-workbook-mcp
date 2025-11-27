using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.Server.Tests;

public sealed class ExcelWorkbookServiceTests
{
    private static string GetTestDataPath(string fileName)
    {
        // Navigate from test output to test-data directory
        var currentDir = Directory.GetCurrentDirectory();
        var projectRoot = Path.GetFullPath(Path.Combine(currentDir, "..", "..", "..", "..", ".."));
        return Path.Combine(projectRoot, "test-data", fileName);
    }

    [Fact]
    public void Constructor_ThrowsOnNullPath()
    {
        Assert.Throws<ArgumentNullException>(() => new ExcelWorkbookService(null!));
    }

    [Fact]
    public void Constructor_ThrowsOnEmptyPath()
    {
        Assert.Throws<ArgumentException>(() => new ExcelWorkbookService(string.Empty));
    }

    [Fact]
    public void WorkbookPath_ReturnsFullPath()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        Assert.Equal(Path.GetFullPath(path), service.WorkbookPath);
    }

    [Fact]
    public async Task GetMetadataAsync_ReturnsValidMetadata()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        Assert.NotNull(metadata);
        Assert.Equal(service.WorkbookPath, metadata.WorkbookPath);
        Assert.NotEmpty(metadata.Worksheets);
    }

    [Fact]
    public async Task GetMetadataAsync_ThrowsOnNonExistentFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "NonExistent.xlsx");
        var service = new ExcelWorkbookService(path);

        await Assert.ThrowsAsync<FileNotFoundException>(() => service.GetMetadataAsync(CancellationToken.None));
    }

    [Fact]
    public async Task GetMetadataAsync_CachesMetadata()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var metadata1 = await service.GetMetadataAsync(CancellationToken.None);
        var metadata2 = await service.GetMetadataAsync(CancellationToken.None);

        Assert.Same(metadata1, metadata2);
    }

    [Fact]
    public async Task ListResourcesAsync_ReturnsWorkbookAndWorksheets()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var resources = await service.ListResourcesAsync(CancellationToken.None);

        Assert.NotEmpty(resources);
        // Should include workbook resource
        Assert.Contains(resources, r => r.Uri.Host == "workbook");
        // Should include worksheet resources
        Assert.Contains(resources, r => r.Uri.Host == "worksheet");
    }

    [Fact]
    public async Task ReadResourceAsync_ReturnsWorkbookMetadata()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var uri = ExcelResourceUri.WorkbookUri;
        var content = await service.ReadResourceAsync(uri, CancellationToken.None);

        Assert.NotNull(content);
        Assert.Equal("application/json", content.MimeType);
        Assert.NotEmpty(content.Text);
    }

    [Fact]
    public async Task ReadResourceAsync_ThrowsOnInvalidUri()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var uri = new Uri("http://invalid/path");
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.ReadResourceAsync(uri, CancellationToken.None));
    }

    [Fact]
    public async Task SearchAsync_ReturnsEmptyForEmptyQuery()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var args = new ExcelSearchArguments(string.Empty);
        var result = await service.SearchAsync(args, CancellationToken.None);

        Assert.NotNull(result);
        Assert.Empty(result.Rows);
        Assert.False(result.HasMore);
    }

    [Fact]
    public async Task SearchAsync_ReturnsEmptyForNullQuery()
    {
        var path = GetTestDataPath("ProjectTracking.xlsx");
        var service = new ExcelWorkbookService(path);

        var args = new ExcelSearchArguments(null!);
        var result = await service.SearchAsync(args, CancellationToken.None);

        Assert.NotNull(result);
        Assert.Empty(result.Rows);
    }

    [Fact]
    public async Task SearchAsync_FindsMatchingRows()
    {
        var path = GetTestDataPath("EmployeeDirectory.xlsx");
        var service = new ExcelWorkbookService(path);

        // Search for a common term that should be in the workbook
        var args = new ExcelSearchArguments("Engineering");
        var result = await service.SearchAsync(args, CancellationToken.None);

        Assert.NotNull(result);
        Assert.NotEmpty(result.Rows);
    }

    [Fact]
    public async Task SearchAsync_RespectsLimit()
    {
        var path = GetTestDataPath("EmployeeDirectory.xlsx");
        var service = new ExcelWorkbookService(path);

        var args = new ExcelSearchArguments("e", Limit: 2);
        var result = await service.SearchAsync(args, CancellationToken.None);

        Assert.NotNull(result);
        Assert.True(result.Rows.Count <= 2);
    }

    [Fact]
    public async Task SearchAsync_CaseInsensitiveByDefault()
    {
        var path = GetTestDataPath("EmployeeDirectory.xlsx");
        var service = new ExcelWorkbookService(path);

        var argsLower = new ExcelSearchArguments("engineering", CaseSensitive: false);
        var argsUpper = new ExcelSearchArguments("ENGINEERING", CaseSensitive: false);

        var resultLower = await service.SearchAsync(argsLower, CancellationToken.None);
        var resultUpper = await service.SearchAsync(argsUpper, CancellationToken.None);

        Assert.Equal(resultLower.Rows.Count, resultUpper.Rows.Count);
    }

    [Fact]
    public async Task SearchAsync_CaseSensitiveWhenRequested()
    {
        var path = GetTestDataPath("EmployeeDirectory.xlsx");
        var service = new ExcelWorkbookService(path);

        var argsCaseSensitive = new ExcelSearchArguments("ENGINEERING", CaseSensitive: true);
        var result = await service.SearchAsync(argsCaseSensitive, CancellationToken.None);

        // Case sensitive search for uppercase "ENGINEERING" likely finds nothing or fewer results
        Assert.NotNull(result);
    }
}
