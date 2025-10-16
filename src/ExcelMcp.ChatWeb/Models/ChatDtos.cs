using System.Text.Json.Nodes;

namespace ExcelMcp.ChatWeb.Models;

public sealed record ChatMessageDto(string Role, string Content);

public sealed record ChatRequestDto(IReadOnlyList<ChatMessageDto> Messages);

public sealed record ToolCallDto(string Name, JsonObject Arguments, string OutputSummary, bool IsError);

public sealed record ChatResponseDto(string Reply, IReadOnlyList<ToolCallDto> ToolCalls);
