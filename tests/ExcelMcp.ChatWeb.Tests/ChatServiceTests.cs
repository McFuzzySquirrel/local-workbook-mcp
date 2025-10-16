using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Services;
using ExcelMcp.Client.Mcp;
using Xunit;

namespace ExcelMcp.ChatWeb.Tests;

public sealed class ChatServiceTests
{
    [Fact]
    public async Task HandleAsync_ProcessesToolCallAndReturnsFinalResponse()
    {
        var ollama = new FakeOllamaClient(
        [
            "{\"type\":\"tool_call\",\"tool\":\"excel-list-structure\",\"arguments\":{}}",
            "{\"type\":\"final_response\",\"message\":\"Here you go.\"}"
        ]);

        var toolDefinition = new McpToolDefinition(
            "excel-list-structure",
            "Summarise workbook layout",
            new JsonObject
            {
                ["type"] = "object"
            });

        var mcp = new FakeMcpClient([toolDefinition])
        {
            ToolResultFactory = (_, _) => Task.FromResult(new McpToolCallResult(
                new List<McpToolContent>
                {
                    new("text", "Sheet1\nSheet2", null)
                },
                false))
        };

        var service = new ChatService(ollama, mcp);
        var request = new ChatRequestDto(new List<ChatMessageDto>
        {
            new("user", "List the sheets in my workbook.")
        });

        var response = await service.HandleAsync(request, CancellationToken.None);

        Assert.Equal("Here you go.", response.Reply);
        Assert.Single(response.ToolCalls);
        var toolCall = response.ToolCalls[0];
        Assert.Equal("excel-list-structure", toolCall.Name);
        Assert.False(toolCall.IsError);
        Assert.Contains("Sheet1", toolCall.OutputSummary);
        Assert.Single(mcp.Invocations);
        Assert.Equal("excel-list-structure", mcp.Invocations[0].Name);

        Assert.True(ollama.Conversations.Count >= 2, "Expected at least two model turns.");
        var followUp = ollama.Conversations[1].Last();
        Assert.Equal("user", followUp.Role);
        Assert.Contains("Tool excel-list-structure", followUp.Content);
    }

    [Fact]
    public async Task HandleAsync_ReturnsPlainTextWhenModelProducesNonJson()
    {
        var ollama = new FakeOllamaClient([
            "Here is what I found in your workbook."
        ]);

        var toolDefinition = new McpToolDefinition("excel-search", "Search workbook", new JsonObject());
        var mcp = new FakeMcpClient([toolDefinition]);
        var service = new ChatService(ollama, mcp);

        var request = new ChatRequestDto(new List<ChatMessageDto>
        {
            new("user", "Tell me about sales.")
        });

        var response = await service.HandleAsync(request, CancellationToken.None);

        Assert.Equal("Here is what I found in your workbook.", response.Reply);
        Assert.Empty(response.ToolCalls);
        Assert.Empty(mcp.Invocations);
    }

    [Fact]
    public async Task HandleAsync_FlagsToolErrorSummary()
    {
        var ollama = new FakeOllamaClient(
        [
            "{\"type\":\"tool_call\",\"tool\":\"excel-preview\",\"arguments\":{\"worksheet\":\"Sheet1\"}}",
            "{\"type\":\"final_response\",\"message\":\"I was unable to preview the sheet because the tool errored.\"}"
        ]);

        var toolDefinition = new McpToolDefinition("excel-preview", "Preview worksheet", new JsonObject());
        var mcp = new FakeMcpClient([toolDefinition])
        {
            ToolResultFactory = (_, _) => Task.FromResult(new McpToolCallResult(
                new List<McpToolContent>(),
                true))
        };

        var service = new ChatService(ollama, mcp);
        var request = new ChatRequestDto(new List<ChatMessageDto>
        {
            new("user", "Show me Sheet1.")
        });

        var response = await service.HandleAsync(request, CancellationToken.None);

        Assert.Single(response.ToolCalls);
        var toolCall = response.ToolCalls[0];
        Assert.True(toolCall.IsError);
        Assert.Contains("Tool reported an error", toolCall.OutputSummary);
        Assert.Equal("I was unable to preview the sheet because the tool errored.", response.Reply);
    }

    private sealed class FakeOllamaClient : IOllamaClient
    {
        private readonly Queue<string> _responses;

        public FakeOllamaClient(IEnumerable<string> responses)
        {
            _responses = new Queue<string>(responses);
        }

        public List<List<OllamaChatMessage>> Conversations { get; } = new();

        public Task<OllamaChatResponse> SendChatAsync(IReadOnlyList<OllamaChatMessage> messages, CancellationToken cancellationToken)
        {
            var snapshot = messages.Select(m => new OllamaChatMessage(m.Role, m.Content)).ToList();
            Conversations.Add(snapshot);

            if (_responses.Count == 0)
            {
                throw new InvalidOperationException("No scripted responses remaining.");
            }

            var next = _responses.Dequeue();
            return Task.FromResult(new OllamaChatResponse(new OllamaChatMessage("assistant", next), true));
        }
    }

    private sealed class FakeMcpClient : IMcpClient
    {
        private readonly IReadOnlyList<McpToolDefinition> _tools;

        public FakeMcpClient(IReadOnlyList<McpToolDefinition> tools)
        {
            _tools = tools;
        }

        public List<(string Name, JsonNode? Arguments)> Invocations { get; } = new();

        public Func<string, JsonNode?, Task<McpToolCallResult>> ToolResultFactory { get; set; } = (_, _) => Task.FromResult(new McpToolCallResult(Array.Empty<McpToolContent>(), false));

        public Task<IReadOnlyList<McpToolDefinition>> ListToolsAsync(CancellationToken cancellationToken)
        {
            return Task.FromResult(_tools);
        }

        public async Task<McpToolCallResult> CallToolAsync(string name, JsonNode? arguments, CancellationToken cancellationToken)
        {
            Invocations.Add((name, arguments));
            return await ToolResultFactory(name, arguments).ConfigureAwait(false);
        }
    }
}
