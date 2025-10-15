using System.Text.Json;

namespace ExcelMcp.Server.Mcp;

internal static class JsonOptions
{
    public static JsonSerializerOptions Serializer { get; } = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = false,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };
}
