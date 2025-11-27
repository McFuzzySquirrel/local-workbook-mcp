using System.Text;
using ClosedXML.Excel;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Mcp;

namespace ExcelMcp.Server.Excel;

public sealed class ExcelWorkbookService
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

                var pivotTables = worksheet.PivotTables
                    .Select(pivot => new PivotTableMetadata(
                        pivot.Name,
                        worksheet.Name,
                        "Pivot Table",
                        pivot.RowLabels.Count(),
                        pivot.ColumnLabels.Count(),
                        pivot.Values.Count()
                    ))
                    .ToArray();

                var columnHeaders = GetWorksheetHeaders(worksheet);
                worksheets.Add(new WorksheetMetadata(worksheet.Name, tables, columnHeaders, pivotTables));
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

                var csv = BuildTableCsv(table, maxRows);
                return new ExcelResourceContent(ExcelResourceUri.CreateTableUri(worksheet.Name, table.Name), "text/csv", csv);
            }

            var worksheetCsv = BuildWorksheetCsv(worksheet, maxRows);
            return new ExcelResourceContent(ExcelResourceUri.CreateWorksheetUri(worksheet.Name), "text/csv", worksheetCsv);
        }, cancellationToken);
    }

    public async Task<ExcelSearchResult> SearchAsync(ExcelSearchArguments arguments, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        if (string.IsNullOrWhiteSpace(arguments.Query))
        {
            return new ExcelSearchResult(Array.Empty<ExcelRowResult>(), false);
        }

        var comparison = arguments.CaseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var limit = arguments.Limit is > 0 ? arguments.Limit.Value : 20;

        return await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(_workbookPath);
            var rows = new List<ExcelRowResult>();
            var query = arguments.Query;
            var hasMore = false;

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

                        ScanTable(table, worksheet.Name, query, comparison, limit, rows, ref hasMore, cancellationToken);
                        if (hasMore)
                        {
                            return new ExcelSearchResult(rows, true);
                        }
                    }
                }
                else
                {
                    ScanWorksheet(worksheet, query, comparison, limit, rows, ref hasMore, cancellationToken);
                    if (hasMore)
                    {
                        return new ExcelSearchResult(rows, true);
                    }
                }
            }

            return new ExcelSearchResult(rows, hasMore);
        }, cancellationToken);
    }

    private static bool WorksheetMatches(string worksheetName, string? requestedWorksheet, StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        if (string.IsNullOrWhiteSpace(requestedWorksheet))
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
        if (string.IsNullOrWhiteSpace(requestedTable))
        {
            return true;
        }

        return string.Equals(tableName, requestedTable, comparison);
    }

    private static bool TableMatches(string tableName, string? requestedTable)
    {
        return TableMatches(tableName, requestedTable, StringComparison.OrdinalIgnoreCase);
    }

    private static void ScanTable(IXLTable table, string worksheetName, string query, StringComparison comparison, int limit, ICollection<ExcelRowResult> results, ref bool hasMore, CancellationToken cancellationToken)
    {
        var headers = table.Fields.Select(f => f.Name).ToArray();
        var rows = table.DataRange.Rows();

        foreach (var row in rows)
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

            results.Add(new ExcelRowResult(worksheetName, table.Name, row.RangeAddress.FirstAddress.RowNumber, values));
            if (results.Count >= limit)
            {
                hasMore = true;
                return;
            }
        }
    }

    private static void ScanWorksheet(IXLWorksheet worksheet, string query, StringComparison comparison, int limit, ICollection<ExcelRowResult> results, ref bool hasMore, CancellationToken cancellationToken)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange is null)
        {
            return;
        }

        var headerRow = usedRange.FirstRowUsed();
        if (headerRow is null)
        {
            return;
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

            results.Add(new ExcelRowResult(worksheet.Name, null, row.RowNumber(), values));
            if (results.Count >= limit)
            {
                hasMore = true;
                return;
            }
        }
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

    private static string BuildTableCsv(IXLTable table, int maxRows)
    {
        var headers = table.Fields.Select(f => f.Name).ToArray();
        var builder = new StringBuilder();
        WriteCsvRow(builder, headers);

        foreach (var row in table.DataRange.Rows().Take(maxRows))
        {
            var values = headers.Select((_, index) => FormatCell(row.Cell(index + 1))).ToArray();
            WriteCsvRow(builder, values);
        }

        return builder.ToString();
    }

    private static string BuildWorksheetCsv(IXLWorksheet worksheet, int maxRows)
    {
        var range = worksheet.RangeUsed();
        if (range is null)
        {
            return string.Empty;
        }

        var builder = new StringBuilder();
        foreach (var row in range.Rows().Take(maxRows))
        {
            var values = row.Cells().Select(FormatCell).ToArray();
            WriteCsvRow(builder, values);
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

    public async Task<PivotTableResult> AnalyzePivotTablesAsync(PivotTableArguments arguments, CancellationToken cancellationToken)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        if (string.IsNullOrWhiteSpace(arguments.Worksheet))
        {
            throw new ArgumentException("Worksheet name is required.", nameof(arguments));
        }

        return await Task.Run(() =>
        {
            using var workbook = new XLWorkbook(_workbookPath);
            var worksheet = workbook.Worksheets.FirstOrDefault(ws => WorksheetMatches(ws.Name, arguments.Worksheet));
            
            if (worksheet is null)
            {
                throw new InvalidOperationException($"Worksheet '{arguments.Worksheet}' not found.");
            }

            var pivotTables = new List<PivotTableInfo>();
            var pivotsToAnalyze = string.IsNullOrWhiteSpace(arguments.PivotTable)
                ? worksheet.PivotTables
                : worksheet.PivotTables.Where(pt => string.Equals(pt.Name, arguments.PivotTable, StringComparison.OrdinalIgnoreCase));

            foreach (var pivot in pivotsToAnalyze)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var rowFields = pivot.RowLabels.Select(field => new PivotFieldInfo(
                    field.CustomName ?? field.SourceName ?? "Unknown",
                    field.SourceName ?? "Unknown",
                    "Row"
                )).ToArray();

                var columnFields = pivot.ColumnLabels.Select(field => new PivotFieldInfo(
                    field.CustomName ?? field.SourceName ?? "Unknown",
                    field.SourceName ?? "Unknown",
                    "Column"
                )).ToArray();

                var dataFields = pivot.Values.Select(field => new PivotFieldInfo(
                    field.CustomName ?? field.SourceName ?? "Unknown",
                    field.SourceName ?? "Unknown",
                    field.SummaryFormula.ToString()
                )).ToArray();

                var filterFields = arguments.IncludeFilters
                    ? pivot.ReportFilters.Select(field => new PivotFieldInfo(
                        field.CustomName ?? field.SourceName ?? "Unknown",
                        field.SourceName ?? "Unknown",
                        "Filter"
                    )).ToArray()
                    : Array.Empty<PivotFieldInfo>();

                var data = ExtractPivotData(pivot, arguments.MaxRows);

                pivotTables.Add(new PivotTableInfo(
                    pivot.Name,
                    worksheet.Name,
                    worksheet.Name,
                    "Pivot Table Range",
                    rowFields,
                    columnFields,
                    dataFields,
                    filterFields,
                    data
                ));
            }

            return new PivotTableResult(pivotTables);
        }, cancellationToken);
    }

    private static IReadOnlyList<PivotDataRow> ExtractPivotData(IXLPivotTable pivot, int maxRows)
    {
        var rows = new List<PivotDataRow>();
        
        var targetCell = pivot.TargetCell;
        if (targetCell is null)
        {
            return rows;
        }
        
        var lastCellUsed = targetCell.Worksheet.LastCellUsed();
        if (lastCellUsed is null)
        {
            return rows;
        }
        
        var range = targetCell.Worksheet.Range(targetCell.Address, lastCellUsed.Address);
        
        var headerRow = range.FirstRow();
        var headers = headerRow.Cells().Select(c => c.GetString()).ToArray();
        
        if (headers.Length == 0)
        {
            return rows;
        }
        
        var dataRows = range.RowsUsed().Skip(1).Take(maxRows);
        
        foreach (var row in dataRows)
        {
            var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < headers.Length && i < row.CellCount(); i++)
            {
                var header = string.IsNullOrWhiteSpace(headers[i]) ? $"Column{i + 1}" : headers[i];
                values[header] = FormatCell(row.Cell(i + 1));
            }
            
            rows.Add(new PivotDataRow(values));
        }
        
        return rows;
    }
}
