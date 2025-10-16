using System.Text.Json.Nodes;

namespace ExcelMcp.Client.Mcp;

public sealed record McpToolDefinition(string Name, string Description, JsonNode InputSchema);

public sealed record McpToolContent(string Type, string? Text, JsonNode? Json);

public sealed record McpToolCallResult(IReadOnlyList<McpToolContent> Content, bool IsError);

public sealed record McpResource(Uri Uri, string Name, string? Description, string? MimeType);

public sealed record McpResourceContent(Uri Uri, string? MimeType, string? Text);
