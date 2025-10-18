using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Models;
using ExcelMcp.Client.Mcp;

namespace ExcelMcp.ChatWeb.Services;

public sealed class ChatService
{
    private readonly ILlmStudioClient _llmClient;
    private readonly IMcpClient _mcpClient;

    public ChatService(ILlmStudioClient llmClient, IMcpClient mcpClient)
    {
        _llmClient = llmClient;
        _mcpClient = mcpClient;
    }

    public async Task<ChatResponseDto> HandleAsync(ChatRequestDto request, CancellationToken cancellationToken)
    {
        var systemPrompt = await BuildSystemPromptAsync(cancellationToken).ConfigureAwait(false);

        var conversation = new List<LlmStudioChatMessage>
        {
            LlmStudioChatMessage.System(systemPrompt)
        };

        foreach (var message in request.Messages)
        {
            conversation.Add(NormalizeMessage(message));
        }

        var toolCalls = new List<ToolCallDto>();
        var fallbackAttempts = 0;
        const int MaxFallbackAttempts = 2;

        while (true)
        {
            LlmStudioChatResponse response;
            try
            {
                response = await _llmClient.SendChatAsync(conversation, cancellationToken).ConfigureAwait(false);
            }
            catch (HttpRequestException ex)
            {
                var message = $"LLM request failed: {ex.Message}";
                return new ChatResponseDto(message, toolCalls);
            }
            catch (Exception ex)
            {
                var message = $"Unexpected error contacting language model: {ex.Message}";
                return new ChatResponseDto(message, toolCalls);
            }
            var rawContent = response.Content.Trim();
            conversation.Add(LlmStudioChatMessage.Assistant(rawContent));

            if (TryParseJsonObject(rawContent, out var payload))
            {
                if (IsToolCall(payload, out var toolName, out var arguments))
                {
                    var result = await _mcpClient.CallToolAsync(toolName, arguments, cancellationToken).ConfigureAwait(false);
                    var summary = SummarizeToolResult(result, out var continuationInstruction);
                    toolCalls.Add(new ToolCallDto(toolName, CloneJson(arguments), summary, result.IsError));

                    var followUp = BuildToolFollowUp(toolName, summary, result.IsError, continuationInstruction);
                    conversation.Add(LlmStudioChatMessage.User(followUp));
                    continue;
                }

                if (IsFinalResponse(payload, out var messageText))
                {
                    return new ChatResponseDto(messageText, toolCalls);
                }
            }

            if (fallbackAttempts < MaxFallbackAttempts)
            {
                fallbackAttempts++;
                var reminder = BuildFormatReminder(toolCalls.Count > 0);
                conversation.Add(LlmStudioChatMessage.User(reminder));
                continue;
            }

            return new ChatResponseDto(rawContent, toolCalls);
        }
    }

    private async Task<string> BuildSystemPromptAsync(CancellationToken cancellationToken)
    {
        var tools = await _mcpClient.ListToolsAsync(cancellationToken).ConfigureAwait(false);
        var resources = await _mcpClient.ListResourcesAsync(cancellationToken).ConfigureAwait(false);
        var worksheetNames = resources
            .Where(static r => string.Equals(r.Uri.Scheme, "excel", StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.Uri.Host, "worksheet", StringComparison.OrdinalIgnoreCase))
            .Select(static r => Uri.UnescapeDataString(r.Uri.AbsolutePath.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty))
            .Where(static name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        var builder = new StringBuilder();
        builder.AppendLine("You are an assistant that helps users understand and work with Excel workbooks.");
        builder.AppendLine("The workbook is already loaded and accessible through the listed tools—never claim you lack access.");
        builder.AppendLine("Use the tools to gather facts before answering. If a user asks about workbook content, call at least one tool first.");
        builder.AppendLine("Do not invent worksheet or table names. When unsure, inspect the structure first.");
        builder.AppendLine("Always reference tools by their exact names (including the 'excel-' prefix).");
        builder.AppendLine();
        builder.AppendLine("Available tools and their schemas:");

        foreach (var tool in tools)
        {
            builder.Append("- ");
            builder.Append(tool.Name);
            if (!string.IsNullOrWhiteSpace(tool.Description))
            {
                builder.Append(": ");
                builder.Append(tool.Description.Trim());
            }

            builder.AppendLine();
            builder.AppendLine($"  Arguments schema: {tool.InputSchema.ToJsonString(new JsonSerializerOptions { WriteIndented = false })}");
        }

        builder.AppendLine();
        builder.AppendLine("Usage tips:");
        builder.AppendLine("- Use excel-list-structure with {} to summarize worksheets, tables, and columns.");
        builder.AppendLine("- Use excel-search when you need to locate rows that match a query. Omit worksheet/table filters if you need to search the entire workbook or you have not confirmed a specific sheet.");
        builder.AppendLine("- Use excel-preview-table to show sample rows from a worksheet or table; pass the returned nextCursor back via the cursor argument to keep paging.");

        if (worksheetNames.Length > 0)
        {
            builder.AppendLine();
            builder.AppendLine("Known worksheets in this workbook:");
            foreach (var worksheet in worksheetNames)
            {
                builder.Append("  - ");
                builder.AppendLine(worksheet);
            }
        }

        builder.AppendLine();
        builder.AppendLine("Respond only with compact JSON—no prose outside JSON.");
        builder.AppendLine("When you need a tool, reply EXACTLY with:");
        builder.AppendLine("{\"type\":\"tool_call\",\"tool\":\"excel-tool-name\",\"arguments\":{...}}");
        builder.AppendLine("When you can answer, reply EXACTLY with:");
        builder.AppendLine("{\"type\":\"final_response\",\"message\":\"concise helpful answer\"}");
        builder.AppendLine("Example tool call: {\"type\":\"tool_call\",\"tool\":\"excel-list-structure\",\"arguments\":{}}");
        builder.AppendLine("Example final response: {\"type\":\"final_response\",\"message\":\"Worksheet summary...\"}");
        builder.AppendLine("Never add code fences or extra commentary. Favour concise answers.");

        return builder.ToString();
    }

    private static LlmStudioChatMessage NormalizeMessage(ChatMessageDto message)
    {
        return message.Role.ToLowerInvariant() switch
        {
            "system" => LlmStudioChatMessage.System(message.Content),
            "assistant" => LlmStudioChatMessage.Assistant(message.Content),
            _ => LlmStudioChatMessage.User(message.Content)
        };
    }

    private static bool TryParseJsonObject(string content, out JsonObject payload)
    {
        payload = null!;
        var candidate = ExtractJsonCandidate(content);
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return false;
        }

        try
        {
            var node = JsonNode.Parse(candidate)?.AsObject();
            if (node is null)
            {
                return false;
            }

            payload = node;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string ExtractJsonCandidate(string content)
    {
        var trimmed = content.Trim();
        if (trimmed.StartsWith("```", StringComparison.Ordinal))
        {
            var endFence = trimmed.LastIndexOf("```", StringComparison.Ordinal);
            if (endFence > 3)
            {
                var firstLineEnd = trimmed.IndexOf('\n');
                if (firstLineEnd > 0)
                {
                    trimmed = trimmed[(firstLineEnd + 1)..endFence];
                }
            }
        }

        return trimmed.Trim();
    }

    private static bool IsToolCall(JsonObject payload, out string toolName, out JsonObject arguments)
    {
        toolName = string.Empty;
        arguments = new JsonObject();

        if (!payload.TryGetPropertyValue("type", out var typeNode))
        {
            return false;
        }

        if (!string.Equals(typeNode?.GetValue<string>(), "tool_call", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!payload.TryGetPropertyValue("tool", out var toolNode) || string.IsNullOrWhiteSpace(toolNode?.GetValue<string>()))
        {
            return false;
        }

        toolName = toolNode!.GetValue<string>();
        if (payload.TryGetPropertyValue("arguments", out var argsNode) && argsNode is JsonObject obj)
        {
            arguments = obj;
        }

        return true;
    }

    private static bool IsFinalResponse(JsonObject payload, out string message)
    {
        message = string.Empty;
        if (!payload.TryGetPropertyValue("type", out var typeNode))
        {
            return false;
        }

        if (!string.Equals(typeNode?.GetValue<string>(), "final_response", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (!payload.TryGetPropertyValue("message", out var messageNode) || string.IsNullOrWhiteSpace(messageNode?.GetValue<string>()))
        {
            return false;
        }

        message = messageNode!.GetValue<string>();
        return true;
    }

    private static string SummarizeToolResult(McpToolCallResult result, out string? continuationInstruction)
    {
        continuationInstruction = null;
        var builder = new StringBuilder();
        var cursorAppended = false;

        foreach (var item in result.Content)
        {
            if (string.Equals(item.Type, "text", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(item.Text))
            {
                builder.AppendLine(item.Text.Trim());
            }
            else if (string.Equals(item.Type, "json", StringComparison.OrdinalIgnoreCase) && item.Json is not null)
            {
                builder.AppendLine(item.Json.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));

                if (item.Json is JsonObject obj)
                {
                    var hasMore = obj.TryGetPropertyValue("hasMore", out var hasMoreNode) && hasMoreNode is JsonValue hasMoreValue && hasMoreValue.TryGetValue<bool>(out var hasMoreFlag) && hasMoreFlag;
                    var nextCursor = obj.TryGetPropertyValue("nextCursor", out var cursorNode) ? cursorNode?.GetValue<string?>() : null;

                    if (hasMore && !string.IsNullOrWhiteSpace(nextCursor))
                    {
                        if (!cursorAppended)
                        {
                            builder.AppendLine($"Next cursor: {nextCursor}");
                            cursorAppended = true;
                        }

                        continuationInstruction ??= $"More results are available. Call the tool again with {{\"cursor\":\"{nextCursor}\"}}.";
                    }
                }
            }
        }

        if (builder.Length == 0)
        {
            builder.Append(result.IsError ? "Tool reported an error with no explanation." : "Tool returned no content.");
        }

        return builder.ToString().Trim();
    }

    private static string BuildFormatReminder(bool toolUsed)
    {
        var reminder = new StringBuilder();
        reminder.AppendLine("Reminder: respond strictly using the JSON formats provided earlier.");
        reminder.AppendLine("If the user requests workbook content or calculations, call an Excel MCP tool (excel-list-structure, excel-search, excel-preview-table) before answering.");
        reminder.AppendLine("Do not guess worksheet/table names—search the entire workbook by omitting filters unless you confirmed a specific sheet.");
        if (!toolUsed)
        {
            reminder.AppendLine("Do not claim you lack access—the workbook tools are available. Start with a tool_call JSON payload now.");
        }
        reminder.Append("Reply with either a tool_call or final_response JSON object only.");
        return reminder.ToString();
    }

    private static string BuildToolFollowUp(string toolName, string toolOutput, bool isError, string? continuationInstruction)
    {
        var builder = new StringBuilder();
        builder.Append("Tool ");
        builder.Append(toolName);
        builder.Append(isError ? " produced an error:" : " returned:");
        builder.AppendLine();
        builder.AppendLine(toolOutput);
        if (!string.IsNullOrWhiteSpace(continuationInstruction))
        {
            builder.AppendLine(continuationInstruction);
        }
        builder.AppendLine("Use this information to continue the conversation and provide the user with a concise, helpful update.");
        return builder.ToString();
    }

    private static JsonObject CloneJson(JsonObject source)
    {
        var clone = JsonNode.Parse(source.ToJsonString())?.AsObject();
        return clone ?? new JsonObject();
    }
}
