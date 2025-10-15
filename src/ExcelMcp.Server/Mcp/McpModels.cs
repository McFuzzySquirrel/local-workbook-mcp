using System.Text.Json.Nodes;

namespace ExcelMcp.Server.Mcp;

internal sealed record McpInitializeResult(
    string ProtocolVersion,
    McpServerInfo ServerInfo,
    McpCapabilities Capabilities
);

internal sealed record McpServerInfo(string Name, string Version);

internal sealed record McpCapabilities(McpToolCapability Tools, McpResourceCapability Resources);

internal sealed record McpToolCapability(bool ListChanged);

internal sealed record McpResourceCapability(bool ListChanged, bool Subscribe = false);

internal sealed record McpToolsListResult(IReadOnlyList<McpToolDefinition> Tools, string? NextCursor);

internal sealed record McpToolDefinition(string Name, string Description, JsonNode InputSchema);

internal sealed record McpToolCallParams(string Name, JsonNode? Arguments);

internal sealed record McpToolCallResult(IReadOnlyList<McpContentItem> Content, bool IsError = false);

internal abstract record McpContentItem(string Type);

internal sealed record McpTextContent(string Text) : McpContentItem("text");

internal sealed record McpJsonContent(JsonNode Json) : McpContentItem("json");

internal sealed record McpResourcesListResult(IReadOnlyList<McpResource> Resources, string? NextCursor);

internal sealed record McpResource(Uri Uri, string Name, string? Description, string? MimeType);

internal sealed record McpResourceContent(Uri Uri, string? MimeType, string? Text, string? Blob = null);

internal sealed record McpResourcesReadResult(IReadOnlyList<McpResourceContent> Contents);

internal sealed record McpErrorResponse(int Code, string Message, object? Data = null);
