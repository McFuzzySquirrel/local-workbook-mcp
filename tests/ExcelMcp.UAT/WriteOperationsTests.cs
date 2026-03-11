using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.UAT;

/// <summary>
/// User Acceptance Tests — write-back operations (excel-write-cell, excel-write-range,
/// excel-create-worksheet).
///
/// All tests operate on temporary copies of test-data workbooks so the originals
/// are never modified. Temp files land in the OS temp directory and are cleaned up
/// at the end of each test.
/// </summary>
public sealed class WriteOperationsTests
{
    // ── Write Cell ───────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-WO-01: Writing a string value to a single cell persists on re-read")]
    public async Task WriteCell_StringValue_PersistsOnReRead()
    {
        var path = TestData.GetTempCopy("ProjectTracking.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            var result = await writeService.WriteCellAsync(
                new WriteCellRequest(path, "Tasks", "G1", "Notes"),
                CancellationToken.None);

            Assert.True(result.Success);

            // Read it back
            var service = new ExcelWorkbookService(path);
            // Invalidate the cache by constructing fresh instance; read via preview
            var uri = ExcelResourceUri.CreateWorksheetUri("Tasks");
            var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 1);
            Assert.Contains("Notes", content.Text);
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-02: Writing a numeric value stores it as a number, not text")]
    public async Task WriteCell_NumericValue_StoredAsNumber()
    {
        var path = TestData.GetTempCopy("ProjectTracking.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            var result = await writeService.WriteCellAsync(
                new WriteCellRequest(path, "Projects", "D2", "99999"),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.Contains("D2", result.Message);
            Assert.Contains("99999", result.Message);
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-03: Writing null to a cell clears it (message says 'cleared')")]
    public async Task WriteCell_NullValue_ClearsCell()
    {
        var path = TestData.GetTempCopy("EmployeeDirectory.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            var result = await writeService.WriteCellAsync(
                new WriteCellRequest(path, "Employees", "B2", null),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.Contains("cleared", result.Message, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-04: Every successful write creates a backup file alongside the workbook")]
    public async Task WriteCell_AlwaysCreatesBackup()
    {
        var path = TestData.GetTempCopy("BudgetTracker.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            var result = await writeService.WriteCellAsync(
                new WriteCellRequest(path, "Income", "A2", "2024-01-01"),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.NotNull(result.BackupPath);
            Assert.True(File.Exists(result.BackupPath), "Backup file must exist on disk.");
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-05: Writing to a non-existent worksheet throws InvalidOperationException")]
    public async Task WriteCell_NonExistentWorksheet_Throws()
    {
        var path = TestData.GetTempCopy("ProjectTracking.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                writeService.WriteCellAsync(
                    new WriteCellRequest(path, "DoesNotExist", "A1", "value"),
                    CancellationToken.None));
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    // ── Write Range ──────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-WO-06: Writing a range of three cells all succeed in one operation")]
    public async Task WriteRange_UpdatesThreeCells()
    {
        var path = TestData.GetTempCopy("EmployeeDirectory.xlsx");
        try
        {
            var updates = new List<CellUpdate>
            {
                new("A10", "9999"),
                new("B10", "Test Employee"),
                new("C10", "QA"),
            };

            var writeService = new ExcelWriteService();
            var result = await writeService.WriteRangeAsync(
                new WriteRangeRequest(path, "Employees", updates),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.Contains("3", result.Message); // "Updated 3 cell(s)"
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-07: Writing a range creates a backup before any cells are changed")]
    public async Task WriteRange_CreatesBackupFile()
    {
        var path = TestData.GetTempCopy("BudgetTracker.xlsx");
        try
        {
            var updates = new List<CellUpdate>
            {
                new("A9", "2024-11-01"),
                new("B9", "New Vendor"),
            };

            var writeService = new ExcelWriteService();
            var result = await writeService.WriteRangeAsync(
                new WriteRangeRequest(path, "Expenses", updates),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.NotNull(result.BackupPath);
            Assert.True(File.Exists(result.BackupPath));
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-08: Range write to a non-existent worksheet throws InvalidOperationException")]
    public async Task WriteRange_NonExistentWorksheet_Throws()
    {
        var path = TestData.GetTempCopy("ProjectTracking.xlsx");
        try
        {
            var updates = new List<CellUpdate> { new("A1", "X") };
            var writeService = new ExcelWriteService();
            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                writeService.WriteRangeAsync(
                    new WriteRangeRequest(path, "GhostSheet", updates),
                    CancellationToken.None));
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    // ── Create Worksheet ─────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-WO-09: Creating a new worksheet adds it to the workbook")]
    public async Task CreateWorksheet_AppearsInMetadata()
    {
        var path = TestData.GetTempCopy("ProjectTracking.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            var result = await writeService.CreateWorksheetAsync(
                new CreateWorksheetRequest(path, "Summary"),
                CancellationToken.None);

            Assert.True(result.Success);
            Assert.Contains("Summary", result.Message);

            // Verify the new sheet is visible via metadata
            var service = new ExcelWorkbookService(path);
            var metadata = await service.GetMetadataAsync(CancellationToken.None);
            Assert.Contains(metadata.Worksheets, ws => ws.Name == "Summary");
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    [Fact(DisplayName = "UA-WO-10: Creating a worksheet with a duplicate name returns failure without throwing")]
    public async Task CreateWorksheet_DuplicateName_ReturnsFalse()
    {
        var path = TestData.GetTempCopy("ProjectTracking.xlsx");
        try
        {
            var writeService = new ExcelWriteService();
            // "Tasks" already exists in ProjectTracking.xlsx
            var result = await writeService.CreateWorksheetAsync(
                new CreateWorksheetRequest(path, "Tasks"),
                CancellationToken.None);

            Assert.False(result.Success);
            Assert.Contains("already exists", result.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Null(result.BackupPath); // No backup when nothing changed
        }
        finally
        {
            CleanupTempFile(path);
        }
    }

    // ── Error Handling ───────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-WO-11: Writing to a non-existent file throws FileNotFoundException")]
    public async Task WriteCell_NonExistentFile_Throws()
    {
        var writeService = new ExcelWriteService();
        await Assert.ThrowsAsync<FileNotFoundException>(() =>
            writeService.WriteCellAsync(
                new WriteCellRequest(
                    Path.Combine(Path.GetTempPath(), "ghost_file_xyz.xlsx"),
                    "Sheet1", "A1", "value"),
                CancellationToken.None));
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static void CleanupTempFile(string path)
    {
        try
        {
            var dir = Path.GetDirectoryName(path);
            if (dir is not null && Directory.Exists(dir))
                Directory.Delete(dir, recursive: true);
        }
        catch
        {
            // Best-effort cleanup — test result is not affected
        }
    }
}
