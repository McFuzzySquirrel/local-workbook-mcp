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
    [Description("Searches for text across all sheets in the workbook")]
    public async Task<string> SearchWorkbook(
        [Description("The text to search for")] string searchText,
        [Description("Maximum number of results to return (default: 100, max: 500)")] int maxResults = 100)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(searchText))
            {
                return CreateErrorResponse("INVALID_INPUT", "Search text cannot be empty");
            }

            if (maxResults < 1 || maxResults > 500)
            {
                return CreateErrorResponse("INVALID_INPUT", "maxResults must be between 1 and 500");
            }

            var arguments = new JsonObject
            {
                ["searchText"] = searchText,
                ["maxResults"] = maxResults
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
    [Description("Searches for text within a specific worksheet")]
    public async Task<string> SearchInSheet(
        [Description("The worksheet name to search in")] string sheetName,
        [Description("The text to search for")] string searchText,
        [Description("Maximum number of results to return (default: 100, max: 500)")] int maxResults = 100)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return CreateErrorResponse("INVALID_INPUT", "Sheet name cannot be empty");
            }

            if (string.IsNullOrWhiteSpace(searchText))
            {
                return CreateErrorResponse("INVALID_INPUT", "Search text cannot be empty");
            }

            if (maxResults < 1 || maxResults > 500)
            {
                return CreateErrorResponse("INVALID_INPUT", "maxResults must be between 1 and 500");
            }

            var arguments = new JsonObject
            {
                ["searchText"] = searchText,
                ["maxResults"] = maxResults
            };

            var result = await _mcpClient.CallToolAsync("excel-search", arguments, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                // Filter results to only include the specified sheet
                var contentString = result.Content.ToString();
                if (string.IsNullOrEmpty(contentString))
                {
                    return CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
                }
                
                var searchResults = JsonSerializer.Deserialize<JsonObject>(contentString);
                if (searchResults != null && searchResults.TryGetPropertyValue("results", out var resultsNode))
                {
                    var results = resultsNode?.AsArray();
                    if (results != null)
                    {
                        var filteredResults = results
                            .Where(r => r?["sheet"]?.ToString().Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true)
                            .Take(maxResults)
                            .ToList();

                        var filtered = new
                        {
                            sheetName = sheetName,
                            searchText = searchText,
                            resultCount = filteredResults.Count,
                            results = filteredResults
                        };

                        return JsonSerializer.Serialize(filtered);
                    }
                }

                return CreateErrorResponse("SHEET_NOT_FOUND", $"No results found in sheet '{sheetName}'");
            }

            return CreateErrorResponse("MCP_ERROR", "Failed to search workbook");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error searching in sheet: {ex.Message}");
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
