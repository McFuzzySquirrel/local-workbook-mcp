using System.Text.Json.Nodes;
using ExcelMcp.Client.Mcp;

namespace ExcelMcp.ChatWeb.Services;

public interface IMcpClient
{
    Task<IReadOnlyList<McpToolDefinition>> ListToolsAsync(CancellationToken cancellationToken);

    Task<McpToolCallResult> CallToolAsync(string name, JsonNode? arguments, CancellationToken cancellationToken);
}
