using ClosedXML.Excel;
using ExcelMcp.Contracts;

namespace ExcelMcp.Server.Excel;

/// <summary>
/// Provides write-back operations against local Excel workbooks via ClosedXML.
/// Every mutating method creates a timestamped backup of the file before writing.
/// The service is intentionally stateless — it opens, modifies, and saves the workbook
/// in a single synchronous Task.Run block to avoid holding a file lock between calls.
/// </summary>
public sealed class ExcelWriteService
{
    /// <summary>
    /// Creates a backup of the workbook, then writes a single cell value.
    /// </summary>
    public Task<WriteResult> WriteCellAsync(WriteCellRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        return Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fullPath = ResolveAndValidate(request.WorkbookPath);
            var backupPath = CreateBackup(fullPath);

            using var workbook = new XLWorkbook(fullPath);
            var worksheet = GetWorksheet(workbook, request.Worksheet);
            var cell = worksheet.Cell(request.CellAddress);

            SetCellValue(cell, request.Value);

            workbook.Save();

            var display = request.Value is null ? "(cleared)" : $"'{request.Value}'";
            return new WriteResult(true,
                $"Cell {request.CellAddress} on '{request.Worksheet}' set to {display}.",
                backupPath);
        }, cancellationToken);
    }

    /// <summary>
    /// Creates a backup of the workbook, then writes multiple cells in one save.
    /// </summary>
    public Task<WriteResult> WriteRangeAsync(WriteRangeRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        return Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fullPath = ResolveAndValidate(request.WorkbookPath);
            var backupPath = CreateBackup(fullPath);

            using var workbook = new XLWorkbook(fullPath);
            var worksheet = GetWorksheet(workbook, request.Worksheet);

            foreach (var update in request.Updates)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = worksheet.Cell(update.CellAddress);
                SetCellValue(cell, update.Value);
            }

            workbook.Save();

            return new WriteResult(true,
                $"Updated {request.Updates.Count} cell(s) on '{request.Worksheet}'.",
                backupPath);
        }, cancellationToken);
    }

    /// <summary>
    /// Creates a backup of the workbook, then adds a new worksheet.
    /// </summary>
    public Task<WriteResult> CreateWorksheetAsync(CreateWorksheetRequest request, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);

        return Task.Run(() =>
        {
            cancellationToken.ThrowIfCancellationRequested();

            var fullPath = ResolveAndValidate(request.WorkbookPath);

            using var workbook = new XLWorkbook(fullPath);

            if (workbook.Worksheets.Any(ws => string.Equals(ws.Name, request.WorksheetName, StringComparison.OrdinalIgnoreCase)))
            {
                return new WriteResult(false,
                    $"Worksheet '{request.WorksheetName}' already exists.");
            }

            var backupPath = CreateBackup(fullPath);
            workbook.AddWorksheet(request.WorksheetName);
            workbook.Save();

            return new WriteResult(true,
                $"Worksheet '{request.WorksheetName}' created.",
                backupPath);
        }, cancellationToken);
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static string ResolveAndValidate(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Workbook path must be provided.", nameof(path));

        var full = Path.GetFullPath(path);
        if (!File.Exists(full))
            throw new FileNotFoundException($"Workbook not found at '{full}'.", full);

        return full;
    }

    private static IXLWorksheet GetWorksheet(XLWorkbook workbook, string name)
    {
        var worksheet = workbook.Worksheets.FirstOrDefault(
            ws => string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase));

        return worksheet ?? throw new InvalidOperationException(
            $"Worksheet '{name}' not found. Available sheets: {string.Join(", ", workbook.Worksheets.Select(ws => ws.Name))}");
    }

    private static void SetCellValue(IXLCell cell, string? value)
    {
        if (value is null)
        {
            cell.Clear();
            return;
        }

        // Try numeric, then boolean, then fall back to string so formulas aren't accidentally broken.
        if (double.TryParse(value, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var numeric))
        {
            cell.Value = numeric;
        }
        else if (bool.TryParse(value, out var boolean))
        {
            cell.Value = boolean;
        }
        else
        {
            cell.Value = value;
        }
    }

    /// <summary>
    /// Copies the workbook to a timestamped sibling file and returns the backup path.
    /// Pattern: <filename>.bak-<yyyyMMdd-HHmmss>.<ext>
    /// </summary>
    private static string CreateBackup(string fullPath)
    {
        var dir = Path.GetDirectoryName(fullPath)!;
        var name = Path.GetFileNameWithoutExtension(fullPath);
        var ext = Path.GetExtension(fullPath);
        var timestamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss");
        var backupPath = Path.Combine(dir, $"{name}.bak-{timestamp}{ext}");
        File.Copy(fullPath, backupPath, overwrite: false);
        return backupPath;
    }
}
