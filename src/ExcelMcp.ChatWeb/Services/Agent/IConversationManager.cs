using ExcelMcp.ChatWeb.Models;
using Microsoft.SemanticKernel.ChatCompletion;

namespace ExcelMcp.ChatWeb.Services.Agent;

/// <summary>
/// Manages conversation history and context window for LLM.
/// </summary>
public interface IConversationManager
{
    /// <summary>
    /// Adds a user query to conversation.
    /// </summary>
    /// <param name="message">User's query text.</param>
    /// <param name="correlationId">Unique ID for tracking.</param>
    void AddUserTurn(string message, string correlationId);

    /// <summary>
    /// Adds an agent response to conversation.
    /// </summary>
    /// <param name="message">Agent's response text.</param>
    /// <param name="correlationId">Matches the query.</param>
    /// <param name="contentType">Type of response.</param>
    void AddAssistantTurn(string message, string correlationId, ContentType contentType);

    /// <summary>
    /// Adds a system notification to conversation (UI display only, not in context window).
    /// </summary>
    /// <param name="message">System message.</param>
    void AddSystemMessage(string message);

    /// <summary>
    /// Retrieves current context window for Semantic Kernel.
    /// </summary>
    /// <returns>SK ChatHistory with last N turns.</returns>
    ChatHistory GetContextForLLM();

    /// <summary>
    /// Retrieves complete conversation for UI display.
    /// </summary>
    /// <returns>All turns including system messages.</returns>
    List<ConversationTurn> GetFullHistory();
}
