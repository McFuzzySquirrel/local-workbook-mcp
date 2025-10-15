using System.Text.Json.Nodes;

namespace ExcelMcp.Client.Mcp;

internal sealed record McpToolDefinition(string Name, string Description, JsonNode InputSchema);

internal abstract record McpContentItem(string Type);

internal sealed record McpTextContent(string Text) : McpContentItem("text");

internal sealed record McpJsonContent(JsonNode Json) : McpContentItem("json");

internal sealed record McpToolCallResult(IReadOnlyList<McpContentItem> Content, bool IsError);

internal sealed record McpResource(Uri Uri, string Name, string? Description, string? MimeType);

internal sealed record McpResourceContent(Uri Uri, string? MimeType, string? Text);
