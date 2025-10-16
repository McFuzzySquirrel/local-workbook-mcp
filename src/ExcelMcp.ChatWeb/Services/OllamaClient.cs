using System.Net.Http.Json;
using System.Text.Json.Serialization;
using ExcelMcp.ChatWeb.Options;
using Microsoft.Extensions.Options;

namespace ExcelMcp.ChatWeb.Services;

public sealed class OllamaClient : IOllamaClient
{
    private readonly HttpClient _httpClient;
    private readonly OllamaOptions _options;

    public OllamaClient(HttpClient httpClient, IOptions<OllamaOptions> options)
    {
        _httpClient = httpClient;
        _options = options.Value;
    }

    public async Task<OllamaChatResponse> SendChatAsync(IReadOnlyList<OllamaChatMessage> messages, CancellationToken cancellationToken)
    {
        var request = new OllamaChatRequest(_options.Model, messages, _options.Temperature);
        using var response = await _httpClient.PostAsJsonAsync("/api/chat", request, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();
        var payload = await response.Content.ReadFromJsonAsync<OllamaChatResponse>(cancellationToken: cancellationToken).ConfigureAwait(false);
        return payload ?? throw new InvalidOperationException("Ollama response payload was empty.");
    }
}

public sealed class OllamaChatRequest
{
    public OllamaChatRequest(string model, IReadOnlyList<OllamaChatMessage> messages, double temperature)
    {
        Model = model;
        Messages = messages;
        Options = new OllamaChatRequestOptions { Temperature = temperature };
    }

    [JsonPropertyName("model")]
    public string Model { get; }

    [JsonPropertyName("messages")]
    public IReadOnlyList<OllamaChatMessage> Messages { get; }

    [JsonPropertyName("stream")]
    public bool Stream => false;

    [JsonPropertyName("options")]
    public OllamaChatRequestOptions Options { get; }
}

public sealed class OllamaChatRequestOptions
{
    [JsonPropertyName("temperature")]
    public double Temperature { get; set; }
}

public sealed record OllamaChatMessage(string Role, string Content)
{
    public static OllamaChatMessage System(string content) => new("system", content);

    public static OllamaChatMessage User(string content) => new("user", content);

    public static OllamaChatMessage Assistant(string content) => new("assistant", content);
}

public sealed record OllamaChatResponse(OllamaChatMessage Message, bool Done)
{
    public string Content => Message.Content;
}
