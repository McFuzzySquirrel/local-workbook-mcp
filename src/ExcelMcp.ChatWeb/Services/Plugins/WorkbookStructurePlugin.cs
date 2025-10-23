using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Services;
using Microsoft.SemanticKernel;

namespace ExcelMcp.ChatWeb.Services.Plugins;

/// <summary>
/// Semantic Kernel plugin providing workbook metadata and structural information.
/// Wraps MCP excel-list-structure tool.
/// </summary>
public class WorkbookStructurePlugin
{
    private readonly IMcpClient _mcpClient;

    public WorkbookStructurePlugin(IMcpClient mcpClient)
    {
        _mcpClient = mcpClient;
    }

    /// <summary>
    /// Lists metadata about sheets and tables (structure ONLY, not actual data).
    /// Use this when the user asks WHAT sheets/tables exist or how many rows/columns.
    /// Do NOT use this when they want to see actual data.
    /// </summary>
    [KernelFunction("list_workbook_structure")]
    [Description("Get metadata about sheets and tables (names, column names, row counts). Use ONLY when user asks 'what sheets exist', 'what tables are there', 'how many rows'. Do NOT use when they want to see actual data - use preview_table instead.")]
    [return: Description("JSON string containing sheets, tables, row/column counts")]
    public async Task<string> ListWorkbookStructure()
    {
        try
        {
            var result = await _mcpClient.CallToolAsync("excel-list-structure", null, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                return result.Content.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
            }

            return CreateErrorResponse("MCP_ERROR", "Failed to retrieve workbook structure");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calling excel-list-structure: {ex.Message}");
        }
    }

    /// <summary>
    /// Gets a simple list of all worksheet names in the workbook.
    /// </summary>
    [KernelFunction("get_sheet_names")]
    [Description("Gets a simple list of all worksheet names in the workbook")]
    [return: Description("Comma-separated list of sheet names")]
    public async Task<string> GetSheetNames()
    {
        try
        {
            var structureJson = await ListWorkbookStructure();
            
            // Check if it's an error response
            if (structureJson.Contains("\"error\":true"))
            {
                return structureJson;
            }

            var structure = JsonSerializer.Deserialize<JsonObject>(structureJson);
            if (structure != null && structure.TryGetPropertyValue("sheets", out var sheetsNode))
            {
                var sheets = sheetsNode?.AsArray();
                if (sheets != null)
                {
                    var sheetNames = sheets
                        .Select(s => s?["name"]?.ToString())
                        .Where(n => n != null)
                        .ToList();

                    return string.Join(", ", sheetNames);
                }
            }

            return CreateErrorResponse("NO_WORKBOOK", "No sheets found in workbook");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error retrieving sheet names: {ex.Message}");
        }
    }

    /// <summary>
    /// Gets information about tables in a specific worksheet.
    /// </summary>
    [KernelFunction("get_table_info")]
    [Description("Gets information about tables in a specific worksheet")]
    public async Task<string> GetTableInfo(
        [Description("The worksheet name to query")] string sheetName)
    {
        try
        {
            var structureJson = await ListWorkbookStructure();
            
            // Check if it's an error response
            if (structureJson.Contains("\"error\":true"))
            {
                return structureJson;
            }

            var structure = JsonSerializer.Deserialize<JsonObject>(structureJson);
            if (structure != null && structure.TryGetPropertyValue("sheets", out var sheetsNode))
            {
                var sheets = sheetsNode?.AsArray();
                if (sheets != null)
                {
                    var targetSheet = sheets.FirstOrDefault(s => 
                        s?["name"]?.ToString().Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);

                    if (targetSheet != null)
                    {
                        var tableInfo = new
                        {
                            sheetName = targetSheet["name"]?.ToString(),
                            tables = targetSheet["tables"]?.AsArray().Select(t => new
                            {
                                name = t?.ToString(),
                                rowCount = targetSheet["rowCount"]?.GetValue<int>(),
                                columnCount = targetSheet["columnCount"]?.GetValue<int>()
                            }).ToList()
                        };

                        return JsonSerializer.Serialize(tableInfo);
                    }
                }
            }

            return CreateErrorResponse("SHEET_NOT_FOUND", $"Sheet '{sheetName}' not found in workbook");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error retrieving table info: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a standardized error response JSON.
    /// </summary>
    private static string CreateErrorResponse(string errorCode, string message)
    {
        var error = new
        {
            error = true,
            errorCode = errorCode,
            message = message,
            timestamp = DateTimeOffset.UtcNow,
            canRetry = errorCode != "SHEET_NOT_FOUND"
        };

        return JsonSerializer.Serialize(error);
    }
}
