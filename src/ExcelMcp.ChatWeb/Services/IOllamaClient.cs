namespace ExcelMcp.ChatWeb.Services;

public interface IOllamaClient
{
    Task<OllamaChatResponse> SendChatAsync(IReadOnlyList<OllamaChatMessage> messages, CancellationToken cancellationToken);
}
