using System.ComponentModel;
using System.Text.Json;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using ModelContextProtocol.Server;

namespace ExcelMcp.Server.Mcp;

/// <summary>
/// MCP tool handlers for Excel workbook operations.
/// Each tool accepts an optional <c>workbook_path</c> parameter.
/// When omitted, the value of the <c>EXCEL_MCP_WORKBOOK</c> environment variable is used,
/// which is set at startup when the server is launched with <c>--workbook &lt;path&gt;</c>.
/// This dual-mode design supports both standalone use (Blazor web app / CLI) and
/// external MCP clients (Claude Desktop, GitHub Copilot, Cursor).
/// </summary>
[McpServerToolType]
internal sealed class ExcelTools
{
    private static string ResolveAndValidatePath(string? path)
    {
        var resolved = !string.IsNullOrWhiteSpace(path)
            ? path
            : Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");

        if (string.IsNullOrWhiteSpace(resolved))
        {
            throw new InvalidOperationException(
                "No workbook path provided. Pass 'workbook_path' in the tool call, " +
                "or set the EXCEL_MCP_WORKBOOK environment variable.");
        }

        var fullPath = Path.GetFullPath(resolved);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Workbook not found at '{fullPath}'.", fullPath);
        }

        return fullPath;
    }

    [McpServerTool(Name = "excel-list-structure")]
    [Description("Summarize worksheets, named tables, column headers, row counts, and pivot tables in the workbook. Call this first to understand what data is available.")]
    public static async Task<string> ListStructureAsync(
        [Description("Path to the Excel workbook (.xlsx or .xls). If omitted, uses the EXCEL_MCP_WORKBOOK environment variable.")] string? workbook_path = null,
        CancellationToken cancellationToken = default)
    {
        var path = ResolveAndValidatePath(workbook_path);
        var service = new ExcelWorkbookService(path);
        var metadata = await service.GetMetadataAsync(cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Serialize(metadata, JsonOptions.Serializer);
    }

    [McpServerTool(Name = "excel-search")]
    [Description("Search the workbook for rows whose cells contain a given text query. Returns matching rows with their worksheet, table, row number, and cell values.")]
    public static async Task<string> SearchAsync(
        [Description("Text to match within cell values.")] string query,
        [Description("Path to the Excel workbook (.xlsx or .xls). If omitted, uses the EXCEL_MCP_WORKBOOK environment variable.")] string? workbook_path = null,
        [Description("Optional worksheet name to restrict the search to a single sheet.")] string? worksheet = null,
        [Description("Optional Excel table name to restrict the search to a single table.")] string? table = null,
        [Description("Maximum number of matching rows to return (1–100). Defaults to 20.")] int limit = 20,
        [Description("Whether to use case-sensitive string matching. Defaults to false.")] bool case_sensitive = false,
        CancellationToken cancellationToken = default)
    {
        var path = ResolveAndValidatePath(workbook_path);
        var service = new ExcelWorkbookService(path);
        var args = new ExcelSearchArguments(query, worksheet, table, limit, case_sensitive);
        var result = await service.SearchAsync(args, cancellationToken).ConfigureAwait(false);

        if (result.Rows.Count == 0)
        {
            return "No matching rows found.";
        }

        return JsonSerializer.Serialize(result, JsonOptions.Serializer);
    }

    [McpServerTool(Name = "excel-preview-table")]
    [Description("Return a CSV preview of a worksheet or named Excel table. Use to inspect raw data before querying.")]
    public static async Task<string> PreviewTableAsync(
        [Description("Worksheet name to preview.")] string worksheet,
        [Description("Path to the Excel workbook (.xlsx or .xls). If omitted, uses the EXCEL_MCP_WORKBOOK environment variable.")] string? workbook_path = null,
        [Description("Optional named Excel table within the worksheet. If omitted, previews the entire worksheet.")] string? table = null,
        [Description("Maximum number of rows to return (1–100). Defaults to 10.")] int rows = 10,
        CancellationToken cancellationToken = default)
    {
        var path = ResolveAndValidatePath(workbook_path);
        var service = new ExcelWorkbookService(path);
        var uri = table is null
            ? ExcelResourceUri.CreateWorksheetUri(worksheet)
            : ExcelResourceUri.CreateTableUri(worksheet, table);
        var content = await service.ReadResourceAsync(uri, cancellationToken, Math.Max(rows, 1)).ConfigureAwait(false);
        return content.Text ?? string.Empty;
    }

    [McpServerTool(Name = "excel-analyze-pivot")]
    [Description("Analyze pivot tables in a worksheet: structure, row/column/data/filter fields, and aggregated data rows.")]
    public static async Task<string> AnalyzePivotAsync(
        [Description("Worksheet name containing the pivot table(s).")] string worksheet,
        [Description("Path to the Excel workbook (.xlsx or .xls). If omitted, uses the EXCEL_MCP_WORKBOOK environment variable.")] string? workbook_path = null,
        [Description("Specific pivot table name to analyze. If omitted, all pivot tables in the worksheet are analyzed.")] string? pivot_table = null,
        [Description("Whether to include report filter fields in the result. Defaults to true.")] bool include_filters = true,
        [Description("Maximum number of data rows to return per pivot table (1–500). Defaults to 100.")] int max_rows = 100,
        CancellationToken cancellationToken = default)
    {
        var path = ResolveAndValidatePath(workbook_path);
        var service = new ExcelWorkbookService(path);
        var args = new PivotTableArguments(worksheet, pivot_table, include_filters, max_rows);
        var result = await service.AnalyzePivotTablesAsync(args, cancellationToken).ConfigureAwait(false);

        if (result.PivotTables.Count == 0)
        {
            return $"No pivot tables found in worksheet '{worksheet}'.";
        }

        return JsonSerializer.Serialize(result, JsonOptions.Serializer);
    }
}
