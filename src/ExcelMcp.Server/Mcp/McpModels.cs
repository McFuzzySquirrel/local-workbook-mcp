using System.Text.Json.Nodes;

namespace ExcelMcp.Server.Mcp;

internal sealed record McpInitializeResult(
    string ProtocolVersion,
    McpServerInfo ServerInfo,
    JsonObject Capabilities
);

internal sealed record McpServerInfo(string Name, string Version);

internal sealed record McpToolsListResult(IReadOnlyList<McpToolDefinition> Tools, string? NextCursor);

internal sealed record McpToolDefinition(string Name, string Description, JsonNode InputSchema);

internal sealed record McpToolCallParams(string Name, JsonNode? Arguments);

internal sealed record McpToolCallResult(IReadOnlyList<McpToolContent> Content, bool IsError = false);

internal sealed record McpToolContent(string Type, string? Text = null, JsonNode? Json = null);

internal sealed record McpResourcesListResult(IReadOnlyList<McpResource> Resources, string? NextCursor);

internal sealed record McpResource(Uri Uri, string Name, string? Description, string? MimeType);

internal sealed record McpResourceContent(Uri Uri, string? MimeType, string? Text, string? Blob = null);

internal sealed record McpResourcesReadResult(IReadOnlyList<McpResourceContent> Contents);

internal sealed record McpErrorResponse(int Code, string Message, object? Data = null);
