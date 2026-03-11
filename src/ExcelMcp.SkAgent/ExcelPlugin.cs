using System.ComponentModel;
using System.Text.Json;
using Microsoft.SemanticKernel;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;

namespace ExcelMcp.SkAgent;

public sealed class ExcelPlugin
{
    private readonly ExcelWorkbookService _service;
    private readonly List<string>? _debugLog;

    public ExcelPlugin(ExcelWorkbookService service, List<string>? debugLog = null)
    {
        _service = service;
        _debugLog = debugLog;
    }

    private void Log(string message)
    {
        _debugLog?.Add(message);
    }

    [KernelFunction("list_structure")]
    [Description("List all worksheets and tables in the workbook with their column headers and row counts")]
    public async Task<string> ListStructureAsync(CancellationToken cancellationToken = default)
    {
        Log("🔧 Tool Called: list_structure");
        var metadata = await _service.GetMetadataAsync(cancellationToken);
        var result = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
        Log($"✅ Returned metadata for {metadata.Worksheets.Count} worksheets");
        return result;
    }

    [KernelFunction("search")]
    [Description("Search for rows containing a query string across worksheets and tables. Case-INSENSITIVE by default. Returns actual matching rows with their values. Use this to find specific data like 'Laptop' or 'North'.")]
    [return: Description("JSON containing matching rows with all their column values")]
    public async Task<string> SearchAsync(
        [Description("The exact text to search for in any cell (case-insensitive)")] string query,
        [Description("Optional: specific worksheet name to limit search")] string? worksheet = null,
        [Description("Optional: specific table name to limit search")] string? table = null,
        [Description("Case sensitive search - set to false for case-insensitive (default: false)")] bool caseSensitive = false,
        [Description("Maximum number of results to return (default: 20)")] int limit = 20,
        CancellationToken cancellationToken = default)
    {
        Log($"🔧 Tool Called: search(query='{query}', worksheet='{worksheet}', table='{table}', caseSensitive={caseSensitive})");
        var args = new ExcelSearchArguments(query, worksheet, table, limit, caseSensitive);

        var result = await _service.SearchAsync(args, cancellationToken);
        Log($"✅ Found {result.Rows.Count} matching rows");
        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    [KernelFunction("preview_table")]
    [Description("Get actual rows from a table or worksheet. Use this to see data, calculate totals, or analyze content. Returns real data in CSV format.")]
    [return: Description("CSV formatted data with headers and row values")]
    public async Task<string> PreviewTableAsync(
        [Description("The worksheet name (e.g., 'Sales', 'Products')")] string worksheet,
        [Description("Optional: the table name within the worksheet")] string? table = null,
        [Description("Number of rows to retrieve (default: 20, use higher for calculations)")] int rows = 20,
        CancellationToken cancellationToken = default)
    {
        Log($"🔧 Tool Called: preview_table(worksheet='{worksheet}', table='{table}', rows={rows})");
        var uri = table != null 
            ? CreateTableUri(worksheet, table)
            : CreateWorksheetUri(worksheet);

        var content = await _service.ReadResourceAsync(uri, cancellationToken, rows);
        var lineCount = content.Text.Split('\n').Length - 1;
        Log($"✅ Returned {lineCount} rows from {worksheet}{(table != null ? $"/{table}" : "")}");
        return content.Text;
    }

    private static Uri CreateWorksheetUri(string worksheetName)
    {
        return new Uri($"excel://worksheet/{Uri.EscapeDataString(worksheetName)}");
    }

    private static Uri CreateTableUri(string worksheetName, string tableName)
    {
        return new Uri($"excel://worksheet/{Uri.EscapeDataString(worksheetName)}/table/{Uri.EscapeDataString(tableName)}");
    }

    [KernelFunction("get_workbook_summary")]
    [Description("Get a summary of the entire workbook including all worksheets, tables, and metadata")]
    public async Task<string> GetWorkbookSummaryAsync(CancellationToken cancellationToken = default)
    {
        Log("🔧 Tool Called: get_workbook_summary");
        var metadata = await _service.GetMetadataAsync(cancellationToken);
        
        var summary = new
        {
            WorkbookPath = metadata.WorkbookPath,
            WorksheetCount = metadata.Worksheets.Count,
            Worksheets = metadata.Worksheets.Select(ws => new
            {
                ws.Name,
                TableCount = ws.Tables.Count,
                ColumnCount = ws.ColumnHeaders.Count,
                Tables = ws.Tables.Select(t => new
                {
                    t.Name,
                    t.RowCount,
                    ColumnCount = t.ColumnHeaders.Count,
                    Columns = t.ColumnHeaders
                }).ToArray()
            }).ToArray(),
            LastLoadedUtc = metadata.LastLoadedUtc
        };

        Log($"✅ Returned summary with {metadata.Worksheets.Count} worksheets");
        return JsonSerializer.Serialize(summary, new JsonSerializerOptions { WriteIndented = true });
    }

    [KernelFunction("write_cell")]
    [Description("Write a value to a single cell in the workbook. A backup is created automatically before saving. Pass null to clear the cell.")]
    [return: Description("JSON WriteResult with success flag, message, and backup path")]
    public async Task<string> WriteCellAsync(
        [Description("Worksheet name (e.g., 'Sales')")] string worksheet,
        [Description("Cell address in A1 notation (e.g. 'B4')")] string cellAddress,
        [Description("Value to write. Numerics stored as numbers, 'true'/'false' as booleans, else text. Null clears the cell.")] string? value = null,
        CancellationToken cancellationToken = default)
    {
        Log($"🔧 Tool Called: write_cell(worksheet='{worksheet}', cellAddress='{cellAddress}', value='{value}')");
        var writeService = new ExcelWriteService();
        var result = await writeService.WriteCellAsync(
            new WriteCellRequest(_service.WorkbookPath, worksheet, cellAddress, value),
            cancellationToken);
        Log($"✅ write_cell result: {result.Message}");
        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    [KernelFunction("write_range")]
    [Description("Write values to multiple cells in one operation. A backup is created automatically before saving.")]
    [return: Description("JSON WriteResult with success flag, message, and backup path")]
    public async Task<string> WriteRangeAsync(
        [Description("Worksheet name (e.g., 'Sales')")] string worksheet,
        [Description("JSON array of {cellAddress, value} objects. E.g. [{\"cellAddress\":\"A1\",\"value\":\"Hello\"}]")] string updatesJson,
        CancellationToken cancellationToken = default)
    {
        Log($"🔧 Tool Called: write_range(worksheet='{worksheet}', updates={updatesJson})");

        List<CellUpdate> updates;
        try
        {
            updates = JsonSerializer.Deserialize<List<CellUpdate>>(updatesJson)
                ?? throw new ArgumentException("Deserialized to null.");
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new WriteResult(false, $"Invalid updatesJson: {ex.Message}"),
                new JsonSerializerOptions { WriteIndented = true });
        }

        var writeService = new ExcelWriteService();
        var result = await writeService.WriteRangeAsync(
            new WriteRangeRequest(_service.WorkbookPath, worksheet, updates),
            cancellationToken);
        Log($"✅ write_range result: {result.Message}");
        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    [KernelFunction("create_worksheet")]
    [Description("Add a new blank worksheet to the workbook. A backup is created automatically before saving.")]
    [return: Description("JSON WriteResult with success flag and message")]
    public async Task<string> CreateWorksheetAsync(
        [Description("Name for the new worksheet")] string worksheetName,
        CancellationToken cancellationToken = default)
    {
        Log($"🔧 Tool Called: create_worksheet(worksheetName='{worksheetName}')");
        var writeService = new ExcelWriteService();
        var result = await writeService.CreateWorksheetAsync(
            new CreateWorksheetRequest(_service.WorkbookPath, worksheetName),
            cancellationToken);
        Log($"✅ create_worksheet result: {result.Message}");
        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}
