namespace ExcelMcp.ChatWeb.Services;

public interface ILlmStudioClient
{
    Task<LlmStudioChatResponse> SendChatAsync(IReadOnlyList<LlmStudioChatMessage> messages, CancellationToken cancellationToken);
}
