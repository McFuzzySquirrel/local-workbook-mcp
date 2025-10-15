using System.Text.Json.Nodes;

namespace ExcelMcp.Client.Mcp;

internal sealed record McpToolDefinition(string Name, string Description, JsonNode InputSchema);

internal sealed record McpToolContent(string Type, string? Text, JsonNode? Json);

internal sealed record McpToolCallResult(IReadOnlyList<McpToolContent> Content, bool IsError);

internal sealed record McpResource(Uri Uri, string Name, string? Description, string? MimeType);

internal sealed record McpResourceContent(Uri Uri, string? MimeType, string? Text);
