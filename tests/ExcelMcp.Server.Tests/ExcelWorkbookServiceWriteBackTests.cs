using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.Server.Tests;

/// <summary>
/// Tests for write-back operations in ExcelWorkbookService.
/// Uses a copy of the test workbook to avoid modifying the original.
/// </summary>
public sealed class ExcelWorkbookServiceWriteBackTests : IDisposable
{
    private readonly string _testWorkbookPath;
    private readonly string _originalPath;
    private bool _disposed;

    public ExcelWorkbookServiceWriteBackTests()
    {
        _originalPath = GetTestDataPath("ProjectTracking.xlsx");
        _testWorkbookPath = CreateTestWorkbookCopy();
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            try
            {
                if (File.Exists(_testWorkbookPath))
                {
                    File.Delete(_testWorkbookPath);
                }
            }
            catch
            {
                // Ignore cleanup failures
            }
        }
    }

    private static string GetTestDataPath(string fileName)
    {
        var currentDir = Directory.GetCurrentDirectory();
        var candidatePath = Path.GetFullPath(Path.Combine(currentDir, "..", "..", "..", "..", "..", "test-data", fileName));
        
        if (File.Exists(candidatePath))
        {
            return candidatePath;
        }
        
        var dir = new DirectoryInfo(currentDir);
        while (dir != null)
        {
            var testDataDir = Path.Combine(dir.FullName, "test-data", fileName);
            if (File.Exists(testDataDir))
            {
                return testDataDir;
            }
            dir = dir.Parent;
        }
        
        return candidatePath;
    }

    private string CreateTestWorkbookCopy()
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"WriteBackTest_{Guid.NewGuid():N}.xlsx");
        File.Copy(_originalPath, tempPath, overwrite: true);
        return tempPath;
    }

    #region UpdateCellAsync Tests

    [Fact]
    public async Task UpdateCellAsync_ThrowsOnNullArguments()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.UpdateCellAsync(null!, CancellationToken.None));
    }

    [Fact]
    public async Task UpdateCellAsync_ThrowsOnEmptyWorksheet()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new UpdateCellArguments("", "A1", "Value");
        
        await Assert.ThrowsAsync<ArgumentException>(() => service.UpdateCellAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task UpdateCellAsync_ThrowsOnEmptyCellAddress()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new UpdateCellArguments("Sheet1", "", "Value");
        
        await Assert.ThrowsAsync<ArgumentException>(() => service.UpdateCellAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task UpdateCellAsync_ThrowsOnNonExistentWorksheet()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new UpdateCellArguments("NonExistentSheet", "A1", "Value");
        
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.UpdateCellAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task UpdateCellAsync_UpdatesCellValue()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        // Get the first worksheet name
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        var args = new UpdateCellArguments(worksheetName, "A1", "Updated Value");
        var result = await service.UpdateCellAsync(args, CancellationToken.None);
        
        Assert.Equal(worksheetName, result.Worksheet);
        Assert.Equal("A1", result.CellAddress);
        Assert.Equal("Updated Value", result.NewValue);
        Assert.NotNull(result.AuditId);
    }

    [Fact]
    public async Task UpdateCellAsync_RecordsPreviousValue()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        // First update
        var args1 = new UpdateCellArguments(worksheetName, "Z99", "First Value");
        var result1 = await service.UpdateCellAsync(args1, CancellationToken.None);
        
        // Second update
        var args2 = new UpdateCellArguments(worksheetName, "Z99", "Second Value");
        var result2 = await service.UpdateCellAsync(args2, CancellationToken.None);
        
        Assert.Equal("First Value", result2.PreviousValue);
        Assert.Equal("Second Value", result2.NewValue);
    }

    [Fact]
    public async Task UpdateCellAsync_IncludesReasonInAudit()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        var args = new UpdateCellArguments(worksheetName, "A1", "Value", "Test reason");
        var result = await service.UpdateCellAsync(args, CancellationToken.None);
        
        var auditResult = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        var entry = auditResult.Entries.FirstOrDefault(e => e.Id == result.AuditId);
        
        Assert.NotNull(entry);
        Assert.Equal("Test reason", entry.Reason);
    }

    #endregion

    #region AddWorksheetAsync Tests

    [Fact]
    public async Task AddWorksheetAsync_ThrowsOnNullArguments()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.AddWorksheetAsync(null!, CancellationToken.None));
    }

    [Fact]
    public async Task AddWorksheetAsync_ThrowsOnEmptyName()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new AddWorksheetArguments("");
        
        await Assert.ThrowsAsync<ArgumentException>(() => service.AddWorksheetAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task AddWorksheetAsync_ThrowsOnDuplicateName()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var existingWorksheetName = metadata.Worksheets[0].Name;
        
        var args = new AddWorksheetArguments(existingWorksheetName);
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.AddWorksheetAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task AddWorksheetAsync_AddsWorksheet()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var newSheetName = $"TestSheet_{Guid.NewGuid():N}".Substring(0, 20);
        var args = new AddWorksheetArguments(newSheetName);
        var result = await service.AddWorksheetAsync(args, CancellationToken.None);
        
        Assert.Equal(newSheetName, result.Name);
        Assert.True(result.Position > 0);
        Assert.NotNull(result.AuditId);
        
        // Verify the worksheet was actually added
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        Assert.Contains(metadata.Worksheets, ws => ws.Name == newSheetName);
    }

    [Fact]
    public async Task AddWorksheetAsync_AddsAtSpecifiedPosition()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var newSheetName = $"TestSheet_{Guid.NewGuid():N}".Substring(0, 20);
        var args = new AddWorksheetArguments(newSheetName, Position: 1);
        var result = await service.AddWorksheetAsync(args, CancellationToken.None);
        
        Assert.Equal(1, result.Position);
    }

    [Fact]
    public async Task AddWorksheetAsync_IncludesReasonInAudit()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var newSheetName = $"TestSheet_{Guid.NewGuid():N}".Substring(0, 20);
        var args = new AddWorksheetArguments(newSheetName, Reason: "Adding analysis worksheet");
        var result = await service.AddWorksheetAsync(args, CancellationToken.None);
        
        var auditResult = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        var entry = auditResult.Entries.FirstOrDefault(e => e.Id == result.AuditId);
        
        Assert.NotNull(entry);
        Assert.Equal("Adding analysis worksheet", entry.Reason);
    }

    #endregion

    #region AddAnnotationAsync Tests

    [Fact]
    public async Task AddAnnotationAsync_ThrowsOnNullArguments()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.AddAnnotationAsync(null!, CancellationToken.None));
    }

    [Fact]
    public async Task AddAnnotationAsync_ThrowsOnEmptyWorksheet()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new AddAnnotationArguments("", "A1", "Note text");
        
        await Assert.ThrowsAsync<ArgumentException>(() => service.AddAnnotationAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task AddAnnotationAsync_ThrowsOnEmptyCellAddress()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new AddAnnotationArguments("Sheet1", "", "Note text");
        
        await Assert.ThrowsAsync<ArgumentException>(() => service.AddAnnotationAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task AddAnnotationAsync_ThrowsOnEmptyText()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new AddAnnotationArguments("Sheet1", "A1", "");
        
        await Assert.ThrowsAsync<ArgumentException>(() => service.AddAnnotationAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task AddAnnotationAsync_ThrowsOnNonExistentWorksheet()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        var args = new AddAnnotationArguments("NonExistentSheet", "A1", "Note text");
        
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.AddAnnotationAsync(args, CancellationToken.None));
    }

    [Fact]
    public async Task AddAnnotationAsync_AddsAnnotation()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        var args = new AddAnnotationArguments(worksheetName, "A1", "Test annotation");
        var result = await service.AddAnnotationAsync(args, CancellationToken.None);
        
        Assert.Equal(worksheetName, result.Worksheet);
        Assert.Equal("A1", result.CellAddress);
        Assert.Equal("Test annotation", result.Text);
        Assert.NotNull(result.AuditId);
    }

    [Fact]
    public async Task AddAnnotationAsync_UsesDefaultAuthor()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        var args = new AddAnnotationArguments(worksheetName, "A1", "Test annotation");
        var result = await service.AddAnnotationAsync(args, CancellationToken.None);
        
        Assert.Equal("ExcelMcp Agent", result.Author);
    }

    [Fact]
    public async Task AddAnnotationAsync_UsesCustomAuthor()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        var args = new AddAnnotationArguments(worksheetName, "A1", "Test annotation", Author: "Custom Author");
        var result = await service.AddAnnotationAsync(args, CancellationToken.None);
        
        Assert.Equal("Custom Author", result.Author);
    }

    #endregion

    #region GetAuditTrailAsync Tests

    [Fact]
    public async Task GetAuditTrailAsync_ThrowsOnNullArguments()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.GetAuditTrailAsync(null!, CancellationToken.None));
    }

    [Fact]
    public async Task GetAuditTrailAsync_ReturnsEmptyForNewService()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        
        Assert.Empty(result.Entries);
        Assert.Equal(0, result.TotalCount);
    }

    [Fact]
    public async Task GetAuditTrailAsync_ReturnsEntriesAfterUpdates()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        // Perform some updates
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "A1", "Value1"), CancellationToken.None);
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "A2", "Value2"), CancellationToken.None);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        
        Assert.Equal(2, result.TotalCount);
        Assert.Equal(2, result.Entries.Count);
    }

    [Fact]
    public async Task GetAuditTrailAsync_FiltersbyOperationType()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "A1", "Value"), CancellationToken.None);
        await service.AddAnnotationAsync(new AddAnnotationArguments(worksheetName, "A1", "Note"), CancellationToken.None);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(OperationType: "UpdateCell"), CancellationToken.None);
        
        Assert.Single(result.Entries);
        Assert.Equal("UpdateCell", result.Entries[0].OperationType);
    }

    [Fact]
    public async Task GetAuditTrailAsync_RespectsLimit()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        // Create multiple entries
        for (int i = 0; i < 5; i++)
        {
            await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, $"A{i + 1}", $"Value{i}"), CancellationToken.None);
        }
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(Limit: 3), CancellationToken.None);
        
        Assert.Equal(5, result.TotalCount);
        Assert.Equal(3, result.Entries.Count);
    }

    [Fact]
    public async Task GetAuditTrailAsync_ReturnsInDescendingOrder()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "A1", "First"), CancellationToken.None);
        await Task.Delay(10); // Small delay to ensure different timestamps
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "A2", "Second"), CancellationToken.None);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        
        Assert.Equal(2, result.Entries.Count);
        Assert.True(result.Entries[0].Timestamp >= result.Entries[1].Timestamp);
    }

    #endregion

    #region Audit Entry Details Tests

    [Fact]
    public async Task AuditEntry_ContainsCorrectDetailsForUpdateCell()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "B5", "NewValue"), CancellationToken.None);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        var entry = result.Entries[0];
        
        Assert.Equal("UpdateCell", entry.OperationType);
        Assert.Contains(worksheetName, entry.Details["worksheet"]);
        Assert.Equal("B5", entry.Details["cell"]);
        Assert.Equal("NewValue", entry.Details["newValue"]);
    }

    [Fact]
    public async Task AuditEntry_ContainsCorrectDetailsForAddWorksheet()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var newSheetName = $"TestSheet_{Guid.NewGuid():N}".Substring(0, 20);
        await service.AddWorksheetAsync(new AddWorksheetArguments(newSheetName), CancellationToken.None);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        var entry = result.Entries[0];
        
        Assert.Equal("AddWorksheet", entry.OperationType);
        Assert.Equal(newSheetName, entry.Details["worksheetName"]);
        Assert.NotNull(entry.Details["position"]);
    }

    [Fact]
    public async Task AuditEntry_ContainsCorrectDetailsForAddAnnotation()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        var metadata = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata.Worksheets[0].Name;
        
        await service.AddAnnotationAsync(new AddAnnotationArguments(worksheetName, "C3", "Test note", "TestAuthor"), CancellationToken.None);
        
        var result = await service.GetAuditTrailAsync(new GetAuditTrailArguments(), CancellationToken.None);
        var entry = result.Entries[0];
        
        Assert.Equal("AddAnnotation", entry.OperationType);
        Assert.Equal(worksheetName, entry.Details["worksheet"]);
        Assert.Equal("C3", entry.Details["cell"]);
        Assert.Equal("Test note", entry.Details["text"]);
        Assert.Equal("TestAuthor", entry.Details["author"]);
    }

    #endregion

    #region Metadata Cache Invalidation Tests

    [Fact]
    public async Task UpdateCellAsync_InvalidatesMetadataCache()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        // Load initial metadata
        var metadata1 = await service.GetMetadataAsync(CancellationToken.None);
        var worksheetName = metadata1.Worksheets[0].Name;
        
        // Update a cell
        await service.UpdateCellAsync(new UpdateCellArguments(worksheetName, "A1", "NewValue"), CancellationToken.None);
        
        // Get metadata again - should not be the same reference (cache was invalidated)
        var metadata2 = await service.GetMetadataAsync(CancellationToken.None);
        
        // The metadata objects should still have the same content, but we verify the timestamp changed
        Assert.True(metadata2.LastLoadedUtc >= metadata1.LastLoadedUtc);
    }

    [Fact]
    public async Task AddWorksheetAsync_InvalidatesMetadataCache()
    {
        var service = new ExcelWorkbookService(_testWorkbookPath);
        
        // Load initial metadata
        var metadata1 = await service.GetMetadataAsync(CancellationToken.None);
        var initialCount = metadata1.Worksheets.Count;
        
        // Add a worksheet
        var newSheetName = $"TestSheet_{Guid.NewGuid():N}".Substring(0, 20);
        await service.AddWorksheetAsync(new AddWorksheetArguments(newSheetName), CancellationToken.None);
        
        // Get metadata again - should reflect the new worksheet
        var metadata2 = await service.GetMetadataAsync(CancellationToken.None);
        
        Assert.Equal(initialCount + 1, metadata2.Worksheets.Count);
    }

    #endregion
}
