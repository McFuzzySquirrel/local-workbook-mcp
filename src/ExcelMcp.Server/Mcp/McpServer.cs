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
            { "excel-preview-table", HandlePreviewAsync }
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
        var summary = new StringBuilder();
        summary.AppendLine($"Workbook: {metadata.WorkbookPath}");
        summary.AppendLine($"Last loaded: {metadata.LastLoadedUtc:u}");
        foreach (var worksheet in metadata.Worksheets)
        {
            summary.AppendLine($"- Worksheet: {worksheet.Name}");
            if (worksheet.ColumnHeaders.Count > 0)
            {
                summary.AppendLine($"  Columns: {string.Join(", ", worksheet.ColumnHeaders)}");
            }

            if (worksheet.Tables.Count == 0)
            {
                continue;
            }

            foreach (var table in worksheet.Tables)
            {
                summary.AppendLine($"  • Table: {table.Name} ({table.RowCount} rows)");
                summary.AppendLine($"    Columns: {string.Join(", ", table.ColumnHeaders)}");
            }
        }

        var content = new McpToolContent("text", Text: summary.ToString().TrimEnd());
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

    private static string GetToolDescription(string toolName)
    {
        return toolName switch
        {
            "excel-list-structure" => "Summarize worksheets, tables, and columns available in the workbook.",
            "excel-search" => "Search the workbook for rows containing a text query across worksheets or tables.",
            "excel-preview-table" => "Return a CSV preview of a worksheet or table.",
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
