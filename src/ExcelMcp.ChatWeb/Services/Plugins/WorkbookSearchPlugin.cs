using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Services;
using Microsoft.SemanticKernel;

namespace ExcelMcp.ChatWeb.Services.Plugins;

/// <summary>
/// Semantic Kernel plugin providing search capabilities across workbook content.
/// Wraps MCP excel-search tool.
/// </summary>
public class WorkbookSearchPlugin
{
    private readonly IMcpClient _mcpClient;

    public WorkbookSearchPlugin(IMcpClient mcpClient)
    {
        _mcpClient = mcpClient;
    }

    /// <summary>
    /// Searches for text across all sheets in the workbook.
    /// </summary>
    [KernelFunction("search_workbook")]
    [Description("Search the workbook for rows containing text query across worksheets or tables")]
    public async Task<string> SearchWorkbook(
        [Description("Text to match within cell values")] string query,
        [Description("Maximum number of matching rows to return (1-100, default: 100)")] int limit = 100,
        [Description("Whether to match using case-sensitive comparison (default: false)")] bool caseSensitive = false)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return CreateErrorResponse("INVALID_INPUT", "Search query cannot be empty");
            }

            if (limit < 1 || limit > 100)
            {
                return CreateErrorResponse("INVALID_INPUT", "limit must be between 1 and 100");
            }

            var arguments = new JsonObject
            {
                ["query"] = query,
                ["limit"] = limit,
                ["caseSensitive"] = caseSensitive
            };

            var result = await _mcpClient.CallToolAsync("excel-search", arguments, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                return result.Content.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
            }

            return CreateErrorResponse("MCP_ERROR", "Failed to search workbook");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calling excel-search: {ex.Message}");
        }
    }

    /// <summary>
    /// Searches for text within a specific worksheet.
    /// </summary>
    [KernelFunction("search_in_sheet")]
    [Description("Search for text within a specific worksheet")]
    public async Task<string> SearchInSheet(
        [Description("Worksheet name to search in")] string worksheet,
        [Description("Text to match within cell values")] string query,
        [Description("Maximum number of matching rows to return (1-100, default: 100)")] int limit = 100,
        [Description("Whether to match using case-sensitive comparison (default: false)")] bool caseSensitive = false)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(worksheet))
            {
                return CreateErrorResponse("INVALID_INPUT", "Worksheet name cannot be empty");
            }

            if (string.IsNullOrWhiteSpace(query))
            {
                return CreateErrorResponse("INVALID_INPUT", "Search query cannot be empty");
            }

            if (limit < 1 || limit > 100)
            {
                return CreateErrorResponse("INVALID_INPUT", "limit must be between 1 and 100");
            }

            var arguments = new JsonObject
            {
                ["query"] = query,
                ["worksheet"] = worksheet,
                ["limit"] = limit,
                ["caseSensitive"] = caseSensitive
            };

            var result = await _mcpClient.CallToolAsync("excel-search", arguments, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                return result.Content.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
            }

            return CreateErrorResponse("MCP_ERROR", "Failed to search in worksheet");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error searching in worksheet: {ex.Message}");
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
            canRetry = errorCode == "MCP_ERROR" || errorCode == "TIMEOUT"
        };

        return JsonSerializer.Serialize(error);
    }
}
