using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Services;
using Microsoft.SemanticKernel;

namespace ExcelMcp.ChatWeb.Services.Plugins;

/// <summary>
/// Semantic Kernel plugin providing write-back capabilities via the MCP server.
/// Wraps excel-write-cell, excel-write-range, and excel-create-worksheet tools.
/// </summary>
public sealed class WorkbookWritePlugin
{
    private readonly IMcpClient _mcpClient;

    public WorkbookWritePlugin(IMcpClient mcpClient)
    {
        _mcpClient = mcpClient;
    }

    [KernelFunction("write_cell")]
    [Description("Write a value to a single cell in the workbook (e.g. 'set B4 to 100', 'update cell A1 to Hello'). A backup is created automatically before saving.")]
    [return: Description("JSON result with success flag, message, and backup path")]
    public async Task<string> WriteCell(
        [Description("Worksheet name (e.g. 'Sales', 'Sheet1')")] string worksheet,
        [Description("Cell address in A1 notation (e.g. 'B4', 'C12')")] string cellAddress,
        [Description("Value to write. Numbers stay numeric, 'true'/'false' become boolean, everything else is text. Leave empty to clear the cell.")] string? value = null)
    {
        try
        {
            var arguments = new JsonObject
            {
                ["worksheet"] = worksheet,
                ["cell_address"] = cellAddress
            };
            if (value is not null)
            {
                arguments["value"] = value;
            }

            var result = await _mcpClient.CallToolAsync("excel-write-cell", arguments, CancellationToken.None);
            return result.Content?.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calling excel-write-cell: {ex.Message}");
        }
    }

    [KernelFunction("write_range")]
    [Description("Write values to multiple cells in one operation (e.g. 'update A1 to Jan, B1 to Feb, C1 to Mar'). A backup is created automatically before saving.")]
    [return: Description("JSON result with success flag, message, and backup path")]
    public async Task<string> WriteRange(
        [Description("Worksheet name (e.g. 'Sales')")] string worksheet,
        [Description("JSON array of cell updates. Each item needs 'cellAddress' (A1 notation) and 'value'. Example: [{\"cellAddress\":\"A1\",\"value\":\"Hello\"},{\"cellAddress\":\"B2\",\"value\":\"42\"}]")] string updatesJson)
    {
        try
        {
            var arguments = new JsonObject
            {
                ["worksheet"] = worksheet,
                ["updates_json"] = updatesJson
            };

            var result = await _mcpClient.CallToolAsync("excel-write-range", arguments, CancellationToken.None);
            return result.Content?.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calling excel-write-range: {ex.Message}");
        }
    }

    [KernelFunction("create_worksheet")]
    [Description("Add a new blank worksheet to the workbook. A backup is created automatically before saving.")]
    [return: Description("JSON result with success flag and message")]
    public async Task<string> CreateWorksheet(
        [Description("Name for the new worksheet")] string worksheetName)
    {
        try
        {
            var arguments = new JsonObject
            {
                ["worksheet_name"] = worksheetName
            };

            var result = await _mcpClient.CallToolAsync("excel-create-worksheet", arguments, CancellationToken.None);
            return result.Content?.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calling excel-create-worksheet: {ex.Message}");
        }
    }

    private static string CreateErrorResponse(string code, string message) =>
        JsonSerializer.Serialize(new { error = code, message });
}
