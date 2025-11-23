using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Services;
using Microsoft.SemanticKernel;

namespace ExcelMcp.ChatWeb.Services.Plugins;

/// <summary>
/// Semantic Kernel plugin providing data retrieval and aggregation capabilities.
/// Wraps MCP excel-preview-table tool and implements aggregations.
/// </summary>
public class DataRetrievalPlugin
{
    private readonly IMcpClient _mcpClient;

    public DataRetrievalPlugin(IMcpClient mcpClient)
    {
        _mcpClient = mcpClient;
    }

    /// <summary>
    /// Previews actual data rows from a worksheet or table.
    /// Use this when the user asks to "show", "display", "get", or "preview" actual DATA or ROWS.
    /// </summary>
    [KernelFunction("preview_table")]
    [Description("Get actual data rows from a worksheet or table. Use when user wants to SEE THE DATA (e.g., 'show me rows', 'display data', 'what's in the table'). Returns CSV data with column headers and cell values.")]
    public async Task<string> PreviewTable(
        [Description("Worksheet name to preview (e.g., 'Sales', 'Inventory')")] string worksheet,
        [Description("Optional table name within the worksheet (e.g., 'SalesTable')")] string? table = null,
        [Description("Maximum number of rows to include (1-100, default: 10)")] int rows = 10)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(worksheet))
            {
                return CreateErrorResponse("INVALID_INPUT", "Worksheet name cannot be empty");
            }

            if (rows < 1 || rows > 100)
            {
                return CreateErrorResponse("INVALID_INPUT", "rows must be between 1 and 100");
            }

            var arguments = new JsonObject
            {
                ["worksheet"] = worksheet,
                ["rows"] = rows
            };

            // Add table parameter if provided
            if (!string.IsNullOrWhiteSpace(table))
            {
                arguments["table"] = table;
            }

            var result = await _mcpClient.CallToolAsync("excel-preview-table", arguments, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                return result.Content.ToString() ?? CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
            }

            return CreateErrorResponse("MCP_ERROR", "Failed to preview table");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calling excel-preview-table: {ex.Message}");
        }
    }

    /// <summary>
    /// Previews data from multiple sheets in a single call.
    /// </summary>
    [KernelFunction("preview_multiple_sheets")]
    [Description("Get data rows from multiple worksheets at once. Useful for cross-sheet comparisons.")]
    public async Task<string> PreviewMultipleSheets(
        [Description("Comma-separated list of worksheet names")] string sheetNames,
        [Description("Maximum number of rows per sheet (default: 5)")] int rowsPerSheet = 5)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetNames))
            {
                return CreateErrorResponse("INVALID_INPUT", "Sheet names cannot be empty");
            }

            var sheets = sheetNames.Split(',').Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
            if (!sheets.Any())
            {
                return CreateErrorResponse("INVALID_INPUT", "No valid sheet names provided");
            }

            if (rowsPerSheet < 1 || rowsPerSheet > 20)
            {
                return CreateErrorResponse("INVALID_INPUT", "rowsPerSheet must be between 1 and 20");
            }

            var results = new Dictionary<string, object>();
            var errors = new List<string>();

            foreach (var sheet in sheets)
            {
                try
                {
                    var arguments = new JsonObject
                    {
                        ["worksheet"] = sheet,
                        ["rows"] = rowsPerSheet
                    };

                    var result = await _mcpClient.CallToolAsync("excel-preview-table", arguments, CancellationToken.None);
                    
                    if (!result.IsError && result.Content != null && result.Content.Count > 0)
                    {
                        var json = result.Content[0].Text;
                        if (!string.IsNullOrEmpty(json))
                        {
                            var node = JsonNode.Parse(json);
                            if (node != null)
                            {
                                results[sheet] = node;
                            }
                            else
                            {
                                results[sheet] = new { error = "Invalid JSON response" };
                            }
                        }
                        else
                        {
                            results[sheet] = new { error = "Empty response" };
                        }
                    }
                    else
                    {
                        results[sheet] = new { error = "Failed to retrieve data" };
                    }
                }
                catch (Exception ex)
                {
                    results[sheet] = new { error = ex.Message };
                    errors.Add($"Error processing sheet '{sheet}': {ex.Message}");
                }
            }

            return JsonSerializer.Serialize(new
            {
                requestedSheets = sheets,
                data = results,
                errors = errors.Any() ? errors : null
            });
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error retrieving multiple sheets: {ex.Message}");
        }
    }

    /// <summary>
    /// Gets data from a specific cell range (e.g., "A1:D10").
    /// </summary>
    [KernelFunction("get_rows_in_range")]
    [Description("Gets data from a specific cell range (e.g., 'A1:D10')")]
    public async Task<string> GetRowsInRange(
        [Description("Sheet name")] string sheetName,
        [Description("Cell range in A1 notation (e.g., 'A1:D10')")] string cellRange)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return CreateErrorResponse("INVALID_INPUT", "Sheet name cannot be empty");
            }

            if (string.IsNullOrWhiteSpace(cellRange))
            {
                return CreateErrorResponse("INVALID_INPUT", "Cell range cannot be empty");
            }

            // Validate cell range format (simple check for A1 notation)
            if (!System.Text.RegularExpressions.Regex.IsMatch(cellRange, @"^[A-Z]+\d+:[A-Z]+\d+$"))
            {
                return CreateErrorResponse("INVALID_RANGE", "Cell range must be in A1 notation (e.g., 'A1:D10')");
            }

            // Parse range to calculate row count
            var rangeParts = cellRange.Split(':');
            var startCell = rangeParts[0];
            var endCell = rangeParts[1];

            var startRow = int.Parse(System.Text.RegularExpressions.Regex.Match(startCell, @"\d+").Value);
            var endRow = int.Parse(System.Text.RegularExpressions.Regex.Match(endCell, @"\d+").Value);
            var rowCount = endRow - startRow + 1;

            if (rowCount > 1000)
            {
                return CreateErrorResponse("RANGE_TOO_LARGE", $"Range spans {rowCount} rows (max: 1000)");
            }

            var arguments = new JsonObject
            {
                ["name"] = sheetName,
                ["rowCount"] = rowCount,
                ["startRow"] = startRow
            };

            var result = await _mcpClient.CallToolAsync("excel-preview-table", arguments, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                // Add range information to response
                var contentString = result.Content.ToString();
                if (string.IsNullOrEmpty(contentString))
                {
                    return CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
                }
                
                var tableData = JsonSerializer.Deserialize<JsonObject>(contentString);
                if (tableData != null)
                {
                    tableData["requestedRange"] = cellRange;
                    return JsonSerializer.Serialize(tableData);
                }

                return contentString;
            }

            return CreateErrorResponse("MCP_ERROR", "Failed to retrieve range data");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error retrieving range: {ex.Message}");
        }
    }

    /// <summary>
    /// Calculates aggregations (sum, average, min, max, count) on a column.
    /// </summary>
    [KernelFunction("calculate_aggregation")]
    [Description("Calculates aggregations (sum, average, min, max, count) on a column")]
    public async Task<string> CalculateAggregation(
        [Description("Table or sheet name")] string name,
        [Description("Column name to aggregate")] string column,
        [Description("Aggregation type: sum, average, min, max, count")] string aggregationType)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return CreateErrorResponse("INVALID_INPUT", "Table or sheet name cannot be empty");
            }

            if (string.IsNullOrWhiteSpace(column))
            {
                return CreateErrorResponse("INVALID_INPUT", "Column name cannot be empty");
            }

            var validAggregations = new[] { "sum", "average", "min", "max", "count" };
            if (!validAggregations.Contains(aggregationType.ToLowerInvariant()))
            {
                return CreateErrorResponse("INVALID_INPUT", $"Invalid aggregation type. Must be one of: {string.Join(", ", validAggregations)}");
            }

            // Retrieve all data from the table (up to 1000 rows)
            var arguments = new JsonObject
            {
                ["name"] = name,
                ["rowCount"] = 1000,
                ["startRow"] = 1
            };

            var result = await _mcpClient.CallToolAsync("excel-preview-table", arguments, CancellationToken.None);
            
            if (!result.IsError && result.Content != null)
            {
                var contentString = result.Content.ToString();
                if (string.IsNullOrEmpty(contentString))
                {
                    return CreateErrorResponse("MCP_ERROR", "Empty response from MCP server");
                }
                
                var tableData = JsonSerializer.Deserialize<JsonObject>(contentString);
                if (tableData != null && 
                    tableData.TryGetPropertyValue("columns", out var columnsNode) &&
                    tableData.TryGetPropertyValue("rows", out var rowsNode))
                {
                    var columns = columnsNode?.AsArray();
                    var rows = rowsNode?.AsArray();

                    if (columns == null || rows == null)
                    {
                        return CreateErrorResponse("MCP_ERROR", "Invalid table data structure");
                    }

                    // Find column index
                    var columnIndex = -1;
                    for (int i = 0; i < columns.Count; i++)
                    {
                        if (columns[i]?.ToString().Equals(column, StringComparison.OrdinalIgnoreCase) == true)
                        {
                            columnIndex = i;
                            break;
                        }
                    }

                    if (columnIndex == -1)
                    {
                        return CreateErrorResponse("COLUMN_NOT_FOUND", $"Column '{column}' not found in table '{name}'");
                    }

                    // Extract column values
                    var values = new List<double>();
                    foreach (var row in rows)
                    {
                        var rowArray = row?.AsArray();
                        if (rowArray != null && columnIndex < rowArray.Count)
                        {
                            var cellValue = rowArray[columnIndex]?.ToString();
                            if (!string.IsNullOrWhiteSpace(cellValue) && double.TryParse(cellValue, out var numericValue))
                            {
                                values.Add(numericValue);
                            }
                        }
                    }

                    if (values.Count == 0 && aggregationType.ToLowerInvariant() != "count")
                    {
                        return CreateErrorResponse("NOT_NUMERIC", $"No numeric values found in column '{column}'");
                    }

                    // Calculate aggregation
                    double aggregationResult = aggregationType.ToLowerInvariant() switch
                    {
                        "sum" => values.Sum(),
                        "average" => values.Average(),
                        "min" => values.Min(),
                        "max" => values.Max(),
                        "count" => values.Count,
                        _ => 0
                    };

                    var response = new
                    {
                        table = name,
                        column = column,
                        aggregationType = aggregationType,
                        result = aggregationResult,
                        valueCount = values.Count,
                        totalRows = rows.Count
                    };

                    return JsonSerializer.Serialize(response);
                }

                return CreateErrorResponse("MCP_ERROR", "Invalid table data structure");
            }

            return CreateErrorResponse("TABLE_NOT_FOUND", $"Table '{name}' not found");
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("MCP_ERROR", $"Error calculating aggregation: {ex.Message}");
        }
    }

    /// <summary>
    /// Filters rows in a worksheet or table based on a column value.
    /// </summary>
    [KernelFunction("filter_table")]
    [Description("Filters rows in a worksheet or table based on a column value. Use for queries like 'Show sales > 1000' or 'Find customers in NY'.")]
    public async Task<string> FilterTable(
        [Description("Worksheet name")] string worksheet,
        [Description("Column name to filter by")] string column,
        [Description("Filter value (for 'between', use two values separated by comma e.g. '10,20')")] string value,
        [Description("Operator: equals, contains, greater_than, less_than, between")] string @operator,
        [Description("Max rows to return (default: 20)")] int limit = 20)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(worksheet)) return CreateErrorResponse("INVALID_INPUT", "Worksheet name required");
            if (string.IsNullOrWhiteSpace(column)) return CreateErrorResponse("INVALID_INPUT", "Column name required");
            
            // Fetch data (limit to 1000 rows for performance)
            var arguments = new JsonObject
            {
                ["worksheet"] = worksheet,
                ["rows"] = 1000
            };

            var result = await _mcpClient.CallToolAsync("excel-preview-table", arguments, CancellationToken.None);
            
            if (result.IsError || result.Content == null || result.Content.Count == 0)
            {
                return CreateErrorResponse("MCP_ERROR", "Failed to retrieve data for filtering");
            }

            var json = result.Content[0].Text;
            if (string.IsNullOrEmpty(json)) return CreateErrorResponse("MCP_ERROR", "Empty response");

            var tableData = JsonNode.Parse(json);
            var columns = tableData?["columns"]?.AsArray();
            var rows = tableData?["rows"]?.AsArray();

            if (columns == null || rows == null) return CreateErrorResponse("MCP_ERROR", "Invalid data structure");

            // Find column index
            int colIndex = -1;
            for (int i = 0; i < columns.Count; i++)
            {
                if (columns[i]?.ToString().Equals(column, StringComparison.OrdinalIgnoreCase) == true)
                {
                    colIndex = i;
                    break;
                }
            }

            if (colIndex == -1) return CreateErrorResponse("COLUMN_NOT_FOUND", $"Column '{column}' not found");

            // Filter rows
            var filteredRows = new List<JsonNode>();
            var op = @operator.ToLowerInvariant();
            
            foreach (var row in rows)
            {
                if (row == null) continue;
                var rowArray = row.AsArray();
                if (rowArray == null || colIndex >= rowArray.Count) continue;

                var cellValue = rowArray[colIndex]?.ToString() ?? "";
                bool match = false;

                try 
                {
                    if (op == "contains")
                    {
                        match = cellValue.Contains(value, StringComparison.OrdinalIgnoreCase);
                    }
                    else if (op == "equals")
                    {
                        match = cellValue.Equals(value, StringComparison.OrdinalIgnoreCase);
                    }
                    else 
                    {
                        // Numeric comparisons
                        if (double.TryParse(cellValue, out double numCell))
                        {
                            if (op == "greater_than" && double.TryParse(value, out double numVal))
                            {
                                match = numCell > numVal;
                            }
                            else if (op == "less_than" && double.TryParse(value, out double numVal2))
                            {
                                match = numCell < numVal2;
                            }
                            else if (op == "between")
                            {
                                var parts = value.Split(',');
                                if (parts.Length == 2 && 
                                    double.TryParse(parts[0], out double min) && 
                                    double.TryParse(parts[1], out double max))
                                {
                                    match = numCell >= min && numCell <= max;
                                }
                            }
                        }
                    }
                }
                catch { /* Ignore parsing errors */ }

                if (match)
                {
                    filteredRows.Add(row);
                    if (filteredRows.Count >= limit) break;
                }
            }

            // Construct response
            var response = new
            {
                columns = columns,
                rows = filteredRows,
                metadata = new 
                { 
                    filter = new { column, @operator, value },
                    matchCount = filteredRows.Count,
                    totalScanned = rows.Count
                }
            };

            return JsonSerializer.Serialize(response);
        }
        catch (Exception ex)
        {
            return CreateErrorResponse("FILTER_ERROR", ex.Message);
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
            canRetry = errorCode == "MCP_ERROR" || errorCode == "TIMEOUT",
            suggestedAction = errorCode switch
            {
                "INVALID_INPUT" => "Check input parameters and try again",
                "INVALID_RANGE" => "Use A1 notation (e.g., 'A1:D10') for cell ranges",
                "RANGE_TOO_LARGE" => "Reduce range to 1000 rows or fewer",
                "COLUMN_NOT_FOUND" => "Verify column name exists in the table",
                "NOT_NUMERIC" => "Ensure column contains numeric values",
                "TABLE_NOT_FOUND" => "Check table or sheet name spelling",
                _ => "Review error message and retry"
            }
        };

        return JsonSerializer.Serialize(error);
    }
}
