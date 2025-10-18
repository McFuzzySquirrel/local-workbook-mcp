using System.Linq;
using System.Net.Http.Json;
using System.Text.Json.Serialization;
using ExcelMcp.ChatWeb.Options;
using Microsoft.Extensions.Options;

namespace ExcelMcp.ChatWeb.Services;

public sealed class LlmStudioClient : ILlmStudioClient
{
    private readonly HttpClient _httpClient;
    private readonly LlmStudioOptions _options;

    public LlmStudioClient(HttpClient httpClient, IOptions<LlmStudioOptions> options)
    {
        _httpClient = httpClient;
        _options = options.Value;
    }

    public async Task<LlmStudioChatResponse> SendChatAsync(IReadOnlyList<LlmStudioChatMessage> messages, CancellationToken cancellationToken)
    {
        var request = new LlmStudioChatRequest(_options.Model, messages, _options.Temperature);
        using var response = await _httpClient.PostAsJsonAsync("/v1/chat/completions", request, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            var detail = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            var message = string.IsNullOrWhiteSpace(detail)
                ? $"LM Studio responded {(int)response.StatusCode} ({response.ReasonPhrase})."
                : $"LM Studio responded {(int)response.StatusCode} ({response.ReasonPhrase}): {detail}";
            throw new HttpRequestException(message);
        }
        var payload = await response.Content.ReadFromJsonAsync<LlmStudioChatResponse>(cancellationToken: cancellationToken).ConfigureAwait(false);
        return payload ?? throw new InvalidOperationException("LLM Studio response payload was empty.");
    }
}

public sealed class LlmStudioChatRequest
{
    public LlmStudioChatRequest(string model, IReadOnlyList<LlmStudioChatMessage> messages, double temperature)
    {
        Model = model;
        Messages = messages;
        Temperature = temperature;
    }

    [JsonPropertyName("model")]
    public string Model { get; }

    [JsonPropertyName("messages")]
    public IReadOnlyList<LlmStudioChatMessage> Messages { get; }

    [JsonPropertyName("temperature")]
    public double Temperature { get; }

    [JsonPropertyName("stream")]
    public bool Stream => false;
}

public sealed record LlmStudioChatMessage(string Role, string Content)
{
    public static LlmStudioChatMessage System(string content) => new("system", content);

    public static LlmStudioChatMessage User(string content) => new("user", content);

    public static LlmStudioChatMessage Assistant(string content) => new("assistant", content);
}

public sealed class LlmStudioChatResponse
{
    [JsonPropertyName("choices")]
    public List<LlmStudioChatChoice> Choices { get; init; } = new();

    public string Content => Choices.FirstOrDefault()?.Message?.Content ?? string.Empty;
}

public sealed class LlmStudioChatChoice
{
    [JsonPropertyName("message")]
    public LlmStudioChatMessage? Message { get; init; }
}
