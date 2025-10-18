using System.Globalization;
using System.Text;
using ClosedXML.Excel;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Mcp;

namespace ExcelMcp.Server.Excel;

internal sealed class ExcelWorkbookService
{
    private const int DefaultPreviewRowCount = 20;

    private readonly string _workbookPath;
    private readonly object _metadataLock = new();

    private WorkbookMetadata? _metadataCache;
    private DateTime _metadataFileTimestamp;

    public ExcelWorkbookService(string workbookPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(workbookPath);
        _workbookPath = Path.GetFullPath(workbookPath);
    }

    public string WorkbookPath => _workbookPath;

    public async Task<WorkbookMetadata> GetMetadataAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var fileInfo = new FileInfo(_workbookPath);
        if (!fileInfo.Exists)
        {
            throw new FileNotFoundException($"Workbook not found at '{_workbookPath}'.", _workbookPath);
        }

        lock (_metadataLock)
        {
            if (_metadataCache is not null && _metadataFileTimestamp == fileInfo.LastWriteTimeUtc)
            {
                return _metadataCache;
            }
        }

        return await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(_workbookPath);
            var worksheets = new List<WorksheetMetadata>();

            foreach (var worksheet in workbook.Worksheets)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var tables = worksheet.Tables
                    .Select(table => new TableMetadata(
                        table.Name,
                        worksheet.Name,
                        table.Fields.Select(f => f.Name).ToArray(),
                        table.DataRange.RowCount()
                    ))
                    .ToArray();

                var columnHeaders = GetWorksheetHeaders(worksheet);
                worksheets.Add(new WorksheetMetadata(worksheet.Name, tables, columnHeaders));
            }

            var metadata = new WorkbookMetadata(
                WorkbookPath,
                worksheets,
                DateTimeOffset.UtcNow
            );

            lock (_metadataLock)
            {
                _metadataCache = metadata;
                _metadataFileTimestamp = fileInfo.LastWriteTimeUtc;
            }

            return metadata;
        }, cancellationToken);
    }

    public async Task<IReadOnlyList<ExcelResourceDescriptor>> ListResourcesAsync(CancellationToken cancellationToken)
    {
        var metadata = await GetMetadataAsync(cancellationToken);
        var descriptors = new List<ExcelResourceDescriptor>
        {
            new(ExcelResourceUri.WorkbookUri, Path.GetFileName(metadata.WorkbookPath), "Workbook summary", "application/json")
        };

        foreach (var worksheet in metadata.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var worksheetUri = ExcelResourceUri.CreateWorksheetUri(worksheet.Name);
            var columnPreview = worksheet.ColumnHeaders.Count > 0
                ? $"Columns: {string.Join(", ", worksheet.ColumnHeaders)}"
                : "No header row detected";
            descriptors.Add(new(worksheetUri, worksheet.Name, columnPreview, "text/csv"));

            foreach (var table in worksheet.Tables)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var tableUri = ExcelResourceUri.CreateTableUri(table.WorksheetName, table.Name);
                var description = $"Table with {table.RowCount} rows";
                descriptors.Add(new(tableUri, $"{table.WorksheetName}::{table.Name}", description, "text/csv"));
            }
        }

        return descriptors;
    }

    public async Task<ExcelPreviewResult> PreviewAsync(ExcelPreviewArguments arguments, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        if (string.IsNullOrWhiteSpace(arguments.Worksheet))
        {
            throw new ArgumentException("Worksheet is required.", nameof(arguments));
        }

        var limit = arguments.Rows is > 0 ? Math.Min(arguments.Rows.Value, 100) : DefaultPreviewRowCount;
        var offset = CursorToken.TryDecode(arguments.Cursor, out var parsedOffset) && parsedOffset > 0 ? parsedOffset : 0;

        return await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(_workbookPath);
            var worksheet = workbook.Worksheets.FirstOrDefault(ws => WorksheetMatches(ws.Name, arguments.Worksheet));
            if (worksheet is null)
            {
                throw new InvalidOperationException($"Worksheet '{arguments.Worksheet}' not found.");
            }

            if (!string.IsNullOrWhiteSpace(arguments.Table))
            {
                var table = worksheet.Tables.FirstOrDefault(t => TableMatches(t.Name, arguments.Table));
                if (table is null)
                {
                    throw new InvalidOperationException($"Table '{arguments.Table}' not found in worksheet '{worksheet.Name}'.");
                }

                return BuildTablePreview(worksheet, table, limit, offset, cancellationToken);
            }

            return BuildWorksheetPreview(worksheet, limit, offset, cancellationToken);
        }, cancellationToken);
    }

    public async Task<ExcelResourceContent> ReadResourceAsync(Uri uri, CancellationToken cancellationToken, int maxRows = DefaultPreviewRowCount)
    {
        if (maxRows <= 0)
        {
            maxRows = DefaultPreviewRowCount;
        }

        if (!ExcelResourceUri.TryParse(uri, out var worksheetName, out var tableName))
        {
            throw new InvalidOperationException($"Unsupported resource URI '{uri}'.");
        }

        if (worksheetName is null)
        {
            var metadata = await GetMetadataAsync(cancellationToken);
            var json = System.Text.Json.JsonSerializer.Serialize(metadata, JsonOptions.Serializer);
            return new ExcelResourceContent(ExcelResourceUri.WorkbookUri, "application/json", json);
        }

        var limit = Math.Max(1, maxRows);

        return await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(_workbookPath);
            var worksheet = workbook.Worksheets.FirstOrDefault(ws => WorksheetMatches(ws.Name, worksheetName));
            if (worksheet is null)
            {
                throw new InvalidOperationException($"Worksheet '{worksheetName}' not found.");
            }

            if (tableName is not null)
            {
                var table = worksheet.Tables.FirstOrDefault(t => TableMatches(t.Name, tableName));
                if (table is null)
                {
                    throw new InvalidOperationException($"Table '{tableName}' not found in worksheet '{worksheetName}'.");
                }

                var preview = BuildTablePreview(worksheet, table, limit, 0, cancellationToken);
                return new ExcelResourceContent(ExcelResourceUri.CreateTableUri(worksheet.Name, table.Name), "text/csv", preview.Csv);
            }

            var worksheetPreview = BuildWorksheetPreview(worksheet, limit, 0, cancellationToken);
            return new ExcelResourceContent(ExcelResourceUri.CreateWorksheetUri(worksheet.Name), "text/csv", worksheetPreview.Csv);
        }, cancellationToken);
    }

    public async Task<ExcelSearchResult> SearchAsync(ExcelSearchArguments arguments, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        if (string.IsNullOrWhiteSpace(arguments.Query))
        {
            return new ExcelSearchResult(Array.Empty<ExcelRowResult>(), false, null);
        }

        var comparison = arguments.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var limit = arguments.Limit is > 0 ? arguments.Limit.Value : 20;
        var offset = CursorToken.TryDecode(arguments.Cursor, out var parsedOffset) ? parsedOffset : 0;

        return await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(_workbookPath);
            var rows = new List<ExcelRowResult>();
            var matchesAfterOffset = 0;
            var remainingSkip = Math.Max(0, offset);
            var hasMore = false;

            foreach (var match in EnumerateMatches(workbook, arguments, comparison, cancellationToken))
            {
                if (remainingSkip > 0)
                {
                    remainingSkip--;
                    continue;
                }

                matchesAfterOffset++;
                if (matchesAfterOffset <= limit)
                {
                    rows.Add(match);
                    continue;
                }

                hasMore = true;
                break;
            }

            var nextCursor = hasMore ? CursorToken.Encode(offset + rows.Count) : null;
            return new ExcelSearchResult(rows, hasMore, nextCursor);
        }, cancellationToken);
    }

    private IEnumerable<ExcelRowResult> EnumerateMatches(XLWorkbook workbook, ExcelSearchArguments arguments, StringComparison comparison, CancellationToken cancellationToken)
    {
        foreach (var worksheet in workbook.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (!WorksheetMatches(worksheet.Name, arguments.Worksheet, comparison))
            {
                continue;
            }

            if (worksheet.Tables.Any())
            {
                foreach (var table in worksheet.Tables)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    if (!TableMatches(table.Name, arguments.Table, comparison))
                    {
                        continue;
                    }

                    foreach (var match in EnumerateTableMatches(table, worksheet.Name, arguments.Query, comparison, cancellationToken))
                    {
                        yield return match;
                    }
                }
            }
            else
            {
                foreach (var match in EnumerateWorksheetMatches(worksheet, arguments.Query, comparison, cancellationToken))
                {
                    yield return match;
                }
            }
        }
    }

    private static IEnumerable<ExcelRowResult> EnumerateTableMatches(IXLTable table, string worksheetName, string query, StringComparison comparison, CancellationToken cancellationToken)
    {
        var headers = table.Fields.Select(f => f.Name).ToArray();

        foreach (var row in table.DataRange.Rows())
        {
            cancellationToken.ThrowIfCancellationRequested();

            var values = new Dictionary<string, string?>(headers.Length, StringComparer.OrdinalIgnoreCase);
            var match = false;

            for (var i = 0; i < headers.Length; i++)
            {
                var cell = row.Cell(i + 1);
                var value = FormatCell(cell);
                values[headers[i]] = value;

                if (!string.IsNullOrEmpty(value) && value.IndexOf(query, comparison) >= 0)
                {
                    match = true;
                }
            }

            if (!match)
            {
                continue;
            }

            yield return new ExcelRowResult(worksheetName, table.Name, row.RangeAddress.FirstAddress.RowNumber, values);
        }
    }

    private static IEnumerable<ExcelRowResult> EnumerateWorksheetMatches(IXLWorksheet worksheet, string query, StringComparison comparison, CancellationToken cancellationToken)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange is null)
        {
            yield break;
        }

        var headerRow = usedRange.FirstRowUsed();
        if (headerRow is null)
        {
            yield break;
        }

        var headers = headerRow.Cells().Select((cell, index) => string.IsNullOrWhiteSpace(cell.GetString()) ? $"Column{index + 1}" : cell.GetString()).ToArray();
        var dataRows = usedRange.RowsUsed().Where(row => row.RowNumber() > headerRow.RowNumber());

        foreach (var row in dataRows)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var values = new Dictionary<string, string?>(headers.Length, StringComparer.OrdinalIgnoreCase);
            var match = false;

            for (var i = 0; i < headers.Length; i++)
            {
                var cell = row.Cell(i + 1);
                var value = FormatCell(cell);
                values[headers[i]] = value;

                if (!string.IsNullOrEmpty(value) && value.IndexOf(query, comparison) >= 0)
                {
                    match = true;
                }
            }

            if (!match)
            {
                continue;
            }

            yield return new ExcelRowResult(worksheet.Name, null, row.RowNumber(), values);
        }
    }

    private static bool WorksheetMatches(string worksheetName, string? requestedWorksheet, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        if (requestedWorksheet is null)
        {
            return true;
        }

        return string.Equals(worksheetName, requestedWorksheet, comparison);
    }

    private static bool WorksheetMatches(string worksheetName, string? requestedWorksheet)
    {
        return WorksheetMatches(worksheetName, requestedWorksheet, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TableMatches(string tableName, string? requestedTable, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        if (requestedTable is null)
        {
            return true;
        }

        return string.Equals(tableName, requestedTable, comparison);
    }

    private static bool TableMatches(string tableName, string? requestedTable)
    {
        return TableMatches(tableName, requestedTable, StringComparison.OrdinalIgnoreCase);
    }

    private static string[] GetWorksheetHeaders(IXLWorksheet worksheet)
    {
        var range = worksheet.RangeUsed();
        if (range is null)
        {
            return Array.Empty<string>();
        }

        var headerRow = range.FirstRowUsed();
        if (headerRow is null)
        {
            return Array.Empty<string>();
        }

        var headers = headerRow.Cells().Select((cell, index) =>
        {
            var header = cell.GetString();
            return string.IsNullOrWhiteSpace(header) ? $"Column{index + 1}" : header;
        });

        return headers.ToArray();
    }

    private ExcelPreviewResult BuildWorksheetPreview(IXLWorksheet worksheet, int limit, int offset, CancellationToken cancellationToken)
    {
        var range = worksheet.RangeUsed();
        if (range is null)
        {
            return new ExcelPreviewResult(worksheet.Name, null, Array.Empty<string>(), Array.Empty<ExcelPreviewRow>(), offset, false, null, string.Empty);
        }

        var headerRow = range.FirstRowUsed();
        if (headerRow is null)
        {
            return new ExcelPreviewResult(worksheet.Name, null, Array.Empty<string>(), Array.Empty<ExcelPreviewRow>(), offset, false, null, string.Empty);
        }

        var headers = headerRow.Cells()
            .Select((cell, index) =>
            {
                var header = cell.GetString();
                return string.IsNullOrWhiteSpace(header) ? $"Column{index + 1}" : header;
            })
            .ToArray();

        if (headers.Length == 0)
        {
            return new ExcelPreviewResult(worksheet.Name, null, Array.Empty<string>(), Array.Empty<ExcelPreviewRow>(), offset, false, null, string.Empty);
        }

        var rows = new List<ExcelPreviewRow>();
        var hasMore = false;
        var skipped = 0;

        foreach (var row in range.RowsUsed())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (row.RowNumber() <= headerRow.RowNumber())
            {
                continue;
            }

            if (skipped < offset)
            {
                skipped++;
                continue;
            }

            if (rows.Count >= limit)
            {
                hasMore = true;
                break;
            }

            var values = new string?[headers.Length];
            for (var i = 0; i < headers.Length; i++)
            {
                values[i] = FormatCell(row.Cell(i + 1));
            }

            rows.Add(new ExcelPreviewRow(row.RowNumber(), values));
        }

        var nextCursor = hasMore ? CursorToken.Encode(offset + rows.Count) : null;
        var csv = BuildCsv(headers, rows);
        return new ExcelPreviewResult(worksheet.Name, null, headers, rows, offset, hasMore, nextCursor, csv);
    }

    private ExcelPreviewResult BuildTablePreview(IXLWorksheet worksheet, IXLTable table, int limit, int offset, CancellationToken cancellationToken)
    {
        var headers = table.Fields.Select(f => f.Name).ToArray();
        if (headers.Length == 0)
        {
            return new ExcelPreviewResult(worksheet.Name, table.Name, Array.Empty<string>(), Array.Empty<ExcelPreviewRow>(), offset, false, null, string.Empty);
        }

        var rows = new List<ExcelPreviewRow>();
        var hasMore = false;
        var skipped = 0;

        foreach (var row in table.DataRange.Rows())
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (skipped < offset)
            {
                skipped++;
                continue;
            }

            if (rows.Count >= limit)
            {
                hasMore = true;
                break;
            }

            var values = new string?[headers.Length];
            for (var i = 0; i < headers.Length; i++)
            {
                values[i] = FormatCell(row.Cell(i + 1));
            }

            rows.Add(new ExcelPreviewRow(row.RangeAddress.FirstAddress.RowNumber, values));
        }

        var nextCursor = hasMore ? CursorToken.Encode(offset + rows.Count) : null;
        var csv = BuildCsv(headers, rows);
        return new ExcelPreviewResult(worksheet.Name, table.Name, headers, rows, offset, hasMore, nextCursor, csv);
    }

    private static string BuildCsv(IReadOnlyList<string> headers, IReadOnlyList<ExcelPreviewRow> rows)
    {
        if (headers.Count == 0 && rows.Count == 0)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();

        if (headers.Count > 0)
        {
            WriteCsvRow(builder, headers.Select(static header => (string?)header).ToArray());
        }

        foreach (var row in rows)
        {
            WriteCsvRow(builder, row.Values);
        }

        return builder.ToString();
    }

    private static void WriteCsvRow(StringBuilder builder, IReadOnlyList<string?> values)
    {
        for (var i = 0; i < values.Count; i++)
        {
            if (i > 0)
            {
                builder.Append(',');
            }

            var value = values[i] ?? string.Empty;
            var needsQuotes = value.Contains(',') || value.Contains('"') || value.Contains('\n') || value.Contains('\r');

            if (needsQuotes)
            {
                builder.Append('"');
                builder.Append(value.Replace("\"", "\"\""));
                builder.Append('"');
            }
            else
            {
                builder.Append(value);
            }
        }

        builder.AppendLine();
    }

    private static string FormatCell(IXLCell cell)
    {
        return cell.GetFormattedString();
    }

    private static class CursorToken
    {
        public static bool TryDecode(string? cursor, out int offset)
        {
            if (string.IsNullOrWhiteSpace(cursor))
            {
                offset = 0;
                return false;
            }

            if (int.TryParse(cursor, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed >= 0)
            {
                offset = parsed;
                return true;
            }

            offset = 0;
            return false;
        }

        public static string Encode(int offset)
        {
            if (offset < 0)
            {
                offset = 0;
            }

            return offset.ToString(CultureInfo.InvariantCulture);
        }
    }
}
