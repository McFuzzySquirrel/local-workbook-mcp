using System.Text.Json;

namespace ExcelMcp.Server.Mcp;

internal sealed class JsonRpcMessage
{
    public JsonRpcMessage(JsonElement root)
    {
        Root = root;
    }

    public JsonElement Root { get; }

    public JsonElement? Id => Root.TryGetProperty("id", out var value) ? value : null;

    public string? Method => Root.TryGetProperty("method", out var method) ? method.GetString() : null;

    public JsonElement? Params => Root.TryGetProperty("params", out var parameters) ? parameters : null;
}
