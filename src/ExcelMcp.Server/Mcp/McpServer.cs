using System.Text;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;

namespace ExcelMcp.Server.Mcp;

internal sealed class McpServer
{
    private readonly ExcelWorkbookService _workbookService;
    private readonly Dictionary<string, Func<JsonNode?, CancellationToken, Task<McpToolCallResult>>> _tools;

    public McpServer(ExcelWorkbookService workbookService)
    {
        _workbookService = workbookService;
        _tools = new Dictionary<string, Func<JsonNode?, CancellationToken, Task<McpToolCallResult>>>(StringComparer.OrdinalIgnoreCase)
        {
            { "excel-list-structure", HandleListStructureAsync },
            { "excel-search", HandleSearchAsync },
            { "excel-preview-table", HandlePreviewAsync },
            { "excel-analyze-pivot", HandleAnalyzePivotAsync },
            { "excel-update-cell", HandleUpdateCellAsync },
            { "excel-add-worksheet", HandleAddWorksheetAsync },
            { "excel-add-annotation", HandleAddAnnotationAsync },
            { "excel-audit-trail", HandleAuditTrailAsync }
        };
    }

    public async Task RunAsync(CancellationToken cancellationToken)
    {
        Log("MCP server starting");
        await using var transport = new JsonRpcTransport(Console.OpenStandardInput(), Console.OpenStandardOutput());
        Log("Transport ready, waiting for messages");
        while (!cancellationToken.IsCancellationRequested)
        {
            JsonRpcMessage? message;
            try
            {
                Log("Awaiting incoming message");
                message = await transport.ReadAsync(cancellationToken).ConfigureAwait(false);
            }
            catch (EndOfStreamException)
            {
                Log("End of stream received, shutting down");
                break;
            }

            if (message is null)
            {
                Log("No message received, shutting down");
                break;
            }

            if (message.Method is null)
            {
                Log("Message without method ignored");
                continue;
            }

            var isNotification = message.Id is null;

            try
            {
                Log($"Dispatching {message.Method}");
                switch (message.Method)
                {
                    case "initialize":
                        Log("Handling initialize");
                        await HandleInitializeAsync(transport, message, cancellationToken).ConfigureAwait(false);
                        Log("Sent initialize result");
                        break;
                    case "initialized":
                        Log("Client reported initialized");
                        // Client completed initialization; no response required.
                        break;
                    case "shutdown":
                        await transport.WriteResultAsync(message.Id, new { }, cancellationToken).ConfigureAwait(false);
                        Log("Sent shutdown ack");
                        break;
                    case "tools/list":
                        Log("Received tools/list");
                        await HandleToolsListAsync(transport, message, cancellationToken).ConfigureAwait(false);
                        Log("Sent tools/list result");
                        break;
                    case "tools/call":
                        await HandleToolsCallAsync(transport, message, cancellationToken).ConfigureAwait(false);
                        Log("Sent tools/call result");
                        break;
                    case "resources/list":
                        await HandleResourcesListAsync(transport, message, cancellationToken).ConfigureAwait(false);
                        Log("Sent resources/list result");
                        break;
                    case "resources/read":
                        await HandleResourcesReadAsync(transport, message, cancellationToken).ConfigureAwait(false);
                        Log("Sent resources/read result");
                        break;
                    case "ping":
                        await transport.WriteResultAsync(message.Id, new { }, cancellationToken).ConfigureAwait(false);
                        Log("Sent ping ack");
                        break;
                    case "exit":
                        Log("Received exit signal");
                        return;
                    default:
                        if (!isNotification)
                        {
                            await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32601, $"Unknown method '{message.Method}'."), cancellationToken).ConfigureAwait(false);
                            Log($"Reported unknown method {message.Method}");
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                if (!isNotification)
                {
                    await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32603, ex.Message), cancellationToken).ConfigureAwait(false);
                    Log($"Error handling {message.Method}: {ex.Message}");
                }
            }
        }
        Log("Server loop ended");
    }

    private async Task HandleInitializeAsync(JsonRpcTransport transport, JsonRpcMessage message, CancellationToken cancellationToken)
    {
        var version = typeof(McpServer).Assembly.GetName().Version?.ToString() ?? "0.1.0";

        var protocolVersion = "2025-06-18";
        if (message.Params is JsonElement paramsElement && paramsElement.ValueKind == JsonValueKind.Object &&
            paramsElement.TryGetProperty("protocolVersion", out var protocolElement) &&
            protocolElement.ValueKind == JsonValueKind.String)
        {
            var requested = protocolElement.GetString();
            if (!string.IsNullOrWhiteSpace(requested))
            {
                protocolVersion = requested!;
            }
        }

        var capabilities = new JsonObject
        {
            ["tools"] = new JsonObject
            {
                ["listChanged"] = false
            },
            ["resources"] = new JsonObject
            {
                ["listChanged"] = false
            }
        };

        var result = new McpInitializeResult(
            protocolVersion,
            new McpServerInfo("excel-local-mcp", version),
            capabilities
        );

        await transport.WriteResultAsync(message.Id, result, cancellationToken).ConfigureAwait(false);
    }

    private async Task HandleToolsListAsync(JsonRpcTransport transport, JsonRpcMessage message, CancellationToken cancellationToken)
    {
        var tools = _tools.Select(tool => new McpToolDefinition(tool.Key, GetToolDescription(tool.Key), BuildInputSchema(tool.Key))).ToArray();
        var result = new McpToolsListResult(tools, NextCursor: null);
        await transport.WriteResultAsync(message.Id, result, cancellationToken).ConfigureAwait(false);
    }

    private async Task HandleToolsCallAsync(JsonRpcTransport transport, JsonRpcMessage message, CancellationToken cancellationToken)
    {
        if (message.Params is null)
        {
            await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32602, "Missing tool invocation parameters."), cancellationToken).ConfigureAwait(false);
            return;
        }

        var callParams = JsonSerializer.Deserialize<McpToolCallParams>(message.Params.Value, JsonOptions.Serializer);
        if (callParams is null || string.IsNullOrWhiteSpace(callParams.Name))
        {
            await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32602, "Invalid tool invocation parameters."), cancellationToken).ConfigureAwait(false);
            return;
        }

        if (!_tools.TryGetValue(callParams.Name, out var handler))
        {
            await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32602, $"Unknown tool '{callParams.Name}'."), cancellationToken).ConfigureAwait(false);
            return;
        }

        var result = await handler(callParams.Arguments, cancellationToken).ConfigureAwait(false);
        await transport.WriteResultAsync(message.Id, result, cancellationToken).ConfigureAwait(false);
    }

    private async Task HandleResourcesListAsync(JsonRpcTransport transport, JsonRpcMessage message, CancellationToken cancellationToken)
    {
        var resources = await _workbookService.ListResourcesAsync(cancellationToken).ConfigureAwait(false);
        var result = new McpResourcesListResult(
            resources.Select(r => new McpResource(r.Uri, r.Name, r.Description, r.MimeType)).ToArray(),
            NextCursor: null
        );
        await transport.WriteResultAsync(message.Id, result, cancellationToken).ConfigureAwait(false);
    }

    private async Task HandleResourcesReadAsync(JsonRpcTransport transport, JsonRpcMessage message, CancellationToken cancellationToken)
    {
        if (message.Params is null || !message.Params.Value.TryGetProperty("uri", out var uriElement))
        {
            await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32602, "Missing resource URI."), cancellationToken).ConfigureAwait(false);
            return;
        }

        var uriText = uriElement.GetString();
        if (string.IsNullOrWhiteSpace(uriText) || !Uri.TryCreate(uriText, UriKind.Absolute, out var uri))
        {
            await transport.WriteErrorAsync(message.Id, new McpErrorResponse(-32602, "Invalid resource URI."), cancellationToken).ConfigureAwait(false);
            return;
        }

        var content = await _workbookService.ReadResourceAsync(uri, cancellationToken).ConfigureAwait(false);
        var result = new McpResourcesReadResult(new[]
        {
            new McpResourceContent(content.Uri, content.MimeType, content.Text)
        });

        await transport.WriteResultAsync(message.Id, result, cancellationToken).ConfigureAwait(false);
    }

    private async Task<McpToolCallResult> HandleListStructureAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        var metadata = await _workbookService.GetMetadataAsync(cancellationToken).ConfigureAwait(false);
        
        // Return the metadata as JSON for programmatic access
        var jsonMetadata = JsonSerializer.Serialize(metadata, JsonOptions.Serializer);
        var content = new McpToolContent("text", Text: jsonMetadata);
        return new McpToolCallResult(new[] { content });
    }

    private async Task<McpToolCallResult> HandleSearchAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        var args = arguments is null
            ? new ExcelSearchArguments(string.Empty)
            : arguments.Deserialize<ExcelSearchArguments>(JsonOptions.Serializer) ?? new ExcelSearchArguments(string.Empty);

        var result = await _workbookService.SearchAsync(args, cancellationToken).ConfigureAwait(false);
        if (result.Rows.Count == 0)
        {
            return new McpToolCallResult(new[]
            {
                new McpToolContent("text", Text: "No matching rows found.")
            });
        }

        var json = new JsonObject
        {
            ["query"] = args.Query,
            ["worksheet"] = args.Worksheet,
            ["table"] = args.Table,
            ["limit"] = args.Limit,
            ["hasMore"] = result.HasMore,
            ["rows"] = new JsonArray(result.Rows.Select(row =>
            {
                var rowJson = new JsonObject
                {
                    ["worksheet"] = row.WorksheetName,
                    ["table"] = row.TableName,
                    ["rowNumber"] = row.RowNumber
                };

                var values = new JsonObject();
                foreach (var pair in row.Values)
                {
                    values[pair.Key] = pair.Value;
                }

                rowJson["values"] = values;
                return rowJson;
            }).ToArray())
        };

        return new McpToolCallResult(new[] { new McpToolContent("json", Json: json) });
    }

    private async Task<McpToolCallResult> HandlePreviewAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        if (arguments is null)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "Worksheet and table arguments are required.") }, true);
        }

        var worksheetName = arguments["worksheet"]?.GetValue<string?>();
        var tableName = arguments["table"]?.GetValue<string?>();
        var rowCount = arguments["rows"]?.GetValue<int?>() ?? 10;

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'worksheet' argument is required.") }, true);
        }

        var uri = tableName is null
            ? ExcelResourceUri.CreateWorksheetUri(worksheetName)
            : ExcelResourceUri.CreateTableUri(worksheetName, tableName);

        try
        {
            var content = await _workbookService.ReadResourceAsync(uri, cancellationToken, Math.Max(rowCount, 1)).ConfigureAwait(false);
            return new McpToolCallResult(new[]
            {
                new McpToolContent("text", Text: content.Text ?? string.Empty)
            });
        }
        catch (Exception ex)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: ex.Message) }, true);
        }
    }

    private async Task<McpToolCallResult> HandleAnalyzePivotAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        if (arguments is null)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "Worksheet argument is required.") }, true);
        }

        var worksheetName = arguments["worksheet"]?.GetValue<string?>();
        var pivotTableName = arguments["pivotTable"]?.GetValue<string?>();
        var includeFilters = arguments["includeFilters"]?.GetValue<bool?>() ?? true;
        var maxRows = arguments["maxRows"]?.GetValue<int?>() ?? 100;

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'worksheet' argument is required.") }, true);
        }

        try
        {
            var args = new PivotTableArguments(worksheetName, pivotTableName, includeFilters, maxRows);
            var result = await _workbookService.AnalyzePivotTablesAsync(args, cancellationToken).ConfigureAwait(false);

            if (result.PivotTables.Count == 0)
            {
                return new McpToolCallResult(new[]
                {
                    new McpToolContent("text", Text: $"No pivot tables found in worksheet '{worksheetName}'.")
                });
            }

            var json = JsonSerializer.Serialize(result, JsonOptions.Serializer);
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: json) });
        }
        catch (Exception ex)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: ex.Message) }, true);
        }
    }

    private async Task<McpToolCallResult> HandleUpdateCellAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        if (arguments is null)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "Worksheet, cell address, and value arguments are required.") }, true);
        }

        var worksheetName = arguments["worksheet"]?.GetValue<string?>();
        var cellAddress = arguments["cell"]?.GetValue<string?>();
        var value = arguments["value"]?.GetValue<string?>();
        var reason = arguments["reason"]?.GetValue<string?>();

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'worksheet' argument is required.") }, true);
        }

        if (string.IsNullOrWhiteSpace(cellAddress))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'cell' argument is required.") }, true);
        }

        if (value is null)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'value' argument is required.") }, true);
        }

        try
        {
            var args = new UpdateCellArguments(worksheetName, cellAddress, value, reason);
            var result = await _workbookService.UpdateCellAsync(args, cancellationToken).ConfigureAwait(false);

            var json = JsonSerializer.Serialize(result, JsonOptions.Serializer);
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: json) });
        }
        catch (Exception ex)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: ex.Message) }, true);
        }
    }

    private async Task<McpToolCallResult> HandleAddWorksheetAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        if (arguments is null)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "Worksheet name argument is required.") }, true);
        }

        var name = arguments["name"]?.GetValue<string?>();
        var position = arguments["position"]?.GetValue<int?>();
        var reason = arguments["reason"]?.GetValue<string?>();

        if (string.IsNullOrWhiteSpace(name))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'name' argument is required.") }, true);
        }

        try
        {
            var args = new AddWorksheetArguments(name, position, reason);
            var result = await _workbookService.AddWorksheetAsync(args, cancellationToken).ConfigureAwait(false);

            var json = JsonSerializer.Serialize(result, JsonOptions.Serializer);
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: json) });
        }
        catch (Exception ex)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: ex.Message) }, true);
        }
    }

    private async Task<McpToolCallResult> HandleAddAnnotationAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        if (arguments is null)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "Worksheet, cell address, and text arguments are required.") }, true);
        }

        var worksheetName = arguments["worksheet"]?.GetValue<string?>();
        var cellAddress = arguments["cell"]?.GetValue<string?>();
        var text = arguments["text"]?.GetValue<string?>();
        var author = arguments["author"]?.GetValue<string?>();

        if (string.IsNullOrWhiteSpace(worksheetName))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'worksheet' argument is required.") }, true);
        }

        if (string.IsNullOrWhiteSpace(cellAddress))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'cell' argument is required.") }, true);
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: "The 'text' argument is required.") }, true);
        }

        try
        {
            var args = new AddAnnotationArguments(worksheetName, cellAddress, text, author);
            var result = await _workbookService.AddAnnotationAsync(args, cancellationToken).ConfigureAwait(false);

            var json = JsonSerializer.Serialize(result, JsonOptions.Serializer);
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: json) });
        }
        catch (Exception ex)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: ex.Message) }, true);
        }
    }

    private async Task<McpToolCallResult> HandleAuditTrailAsync(JsonNode? arguments, CancellationToken cancellationToken)
    {
        DateTimeOffset? since = null;
        DateTimeOffset? until = null;
        string? operationType = null;
        int? limit = null;

        if (arguments is not null)
        {
            var sinceStr = arguments["since"]?.GetValue<string?>();
            if (!string.IsNullOrWhiteSpace(sinceStr) && DateTimeOffset.TryParse(sinceStr, out var parsedSince))
            {
                since = parsedSince;
            }

            var untilStr = arguments["until"]?.GetValue<string?>();
            if (!string.IsNullOrWhiteSpace(untilStr) && DateTimeOffset.TryParse(untilStr, out var parsedUntil))
            {
                until = parsedUntil;
            }

            operationType = arguments["operationType"]?.GetValue<string?>();
            limit = arguments["limit"]?.GetValue<int?>();
        }

        try
        {
            var args = new GetAuditTrailArguments(since, until, operationType, limit);
            var result = await _workbookService.GetAuditTrailAsync(args, cancellationToken).ConfigureAwait(false);

            var json = JsonSerializer.Serialize(result, JsonOptions.Serializer);
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: json) });
        }
        catch (Exception ex)
        {
            return new McpToolCallResult(new[] { new McpToolContent("text", Text: ex.Message) }, true);
        }
    }

    private static string GetToolDescription(string toolName)
    {
        return toolName switch
        {
            "excel-list-structure" => "Summarize worksheets, tables, and columns available in the workbook.",
            "excel-search" => "Search the workbook for rows containing a text query across worksheets or tables.",
            "excel-preview-table" => "Return a CSV preview of a worksheet or table.",
            "excel-analyze-pivot" => "Analyze pivot tables in a worksheet, including structure, fields, and aggregated data.",
            "excel-update-cell" => "Update the value of a cell in a worksheet. Use for making corrections or adding data.",
            "excel-add-worksheet" => "Add a new worksheet to the workbook. Use for organizing new data or analysis.",
            "excel-add-annotation" => "Add an annotation (comment) to a cell. Use for documenting findings or notes.",
            "excel-audit-trail" => "Get the audit trail of changes made to the workbook during this session.",
            _ => "Excel tool"
        };
    }

    private static JsonNode BuildInputSchema(string toolName)
    {
        return toolName switch
        {
            "excel-list-structure" => JsonNode.Parse("{ \"type\": \"object\", \"properties\": {} }")!,
            "excel-search" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Text to match within cell values."},
                    "worksheet": {"type": "string", "description": "Optional worksheet name filter."},
                    "table": {"type": "string", "description": "Optional Excel table name filter."},
                    "limit": {"type": "integer", "minimum": 1, "maximum": 100, "description": "Maximum number of matching rows."},
                    "caseSensitive": {"type": "boolean", "description": "Whether to match using case-sensitive comparison."}
                },
                "required": ["query"]
            }
            """)!,
            "excel-preview-table" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "worksheet": {"type": "string", "description": "Worksheet to preview."},
                    "table": {"type": "string", "description": "Optional table within the worksheet."},
                    "rows": {"type": "integer", "minimum": 1, "maximum": 100, "description": "Maximum number of rows to include."}
                },
                "required": ["worksheet"]
            }
            """)!,
            "excel-analyze-pivot" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "worksheet": {"type": "string", "description": "Worksheet containing the pivot table."},
                    "pivotTable": {"type": "string", "description": "Optional specific pivot table name. If omitted, all pivot tables in the worksheet are analyzed."},
                    "includeFilters": {"type": "boolean", "description": "Whether to include filter fields in the analysis."},
                    "maxRows": {"type": "integer", "minimum": 1, "maximum": 1000, "description": "Maximum number of data rows to include from the pivot table."}
                },
                "required": ["worksheet"]
            }
            """)!,
            "excel-update-cell" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "worksheet": {"type": "string", "description": "Name of the worksheet containing the cell."},
                    "cell": {"type": "string", "description": "Cell address to update (e.g., 'A1', 'B5')."},
                    "value": {"type": "string", "description": "New value to set in the cell."},
                    "reason": {"type": "string", "description": "Optional reason for the update (for audit trail)."}
                },
                "required": ["worksheet", "cell", "value"]
            }
            """)!,
            "excel-add-worksheet" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Name for the new worksheet."},
                    "position": {"type": "integer", "minimum": 1, "description": "Optional position for the worksheet (1-based)."},
                    "reason": {"type": "string", "description": "Optional reason for adding the worksheet (for audit trail)."}
                },
                "required": ["name"]
            }
            """)!,
            "excel-add-annotation" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "worksheet": {"type": "string", "description": "Name of the worksheet containing the cell."},
                    "cell": {"type": "string", "description": "Cell address to annotate (e.g., 'A1', 'B5')."},
                    "text": {"type": "string", "description": "Annotation text (comment) to add."},
                    "author": {"type": "string", "description": "Optional author name for the annotation."}
                },
                "required": ["worksheet", "cell", "text"]
            }
            """)!,
            "excel-audit-trail" => JsonNode.Parse("""
            {
                "type": "object",
                "properties": {
                    "since": {"type": "string", "format": "date-time", "description": "Optional start timestamp for filtering (ISO 8601)."},
                    "until": {"type": "string", "format": "date-time", "description": "Optional end timestamp for filtering (ISO 8601)."},
                    "operationType": {"type": "string", "description": "Optional filter by operation type (UpdateCell, AddWorksheet, AddAnnotation)."},
                    "limit": {"type": "integer", "minimum": 1, "maximum": 1000, "description": "Maximum number of entries to return."}
                }
            }
            """)!,
            _ => JsonNode.Parse("{ }")!
        };
    }

    private static void Log(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return;
        }

        try
        {
            Console.Error.WriteLine($"[{DateTimeOffset.UtcNow:O}] {message}");
        }
        catch
        {
            // ignore logging failures
        }
    }
}
