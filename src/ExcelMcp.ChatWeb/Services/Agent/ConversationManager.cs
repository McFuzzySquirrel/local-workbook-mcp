using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Options;
using Microsoft.Extensions.Options;
using Microsoft.SemanticKernel.ChatCompletion;

namespace ExcelMcp.ChatWeb.Services.Agent;

/// <summary>
/// Manages conversation history and context window for the agent.
/// Maintains a rolling window of the last N turns for LLM context.
/// Uses the WorkbookSession to store conversation history.
/// </summary>
public class ConversationManager : IConversationManager
{
    private readonly ConversationOptions _options;
    private readonly WorkbookSession _session;

    public ConversationManager(IOptions<ConversationOptions> options, WorkbookSession session)
    {
        _options = options?.Value ?? throw new ArgumentNullException(nameof(options));
        _session = session ?? throw new ArgumentNullException(nameof(session));
    }

    /// <summary>
    /// Adds a user message to both full history and context window.
    /// </summary>
    public void AddUserTurn(string message, string correlationId)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            throw new ArgumentException("Message cannot be empty", nameof(message));
        }

        var turn = new ConversationTurn
        {
            Role = "user",
            Content = message,
            CorrelationId = correlationId,
            Timestamp = DateTimeOffset.UtcNow
        };

        _session.ConversationHistory.Add(turn);
        _session.ContextWindow.AddUserMessage(message);
        _session.UpdateActivity();

        // Evict oldest turns if context window exceeds limit
        EvictOldTurnsIfNeeded();
    }

    /// <summary>
    /// Adds an assistant message to both full history and context window.
    /// </summary>
    public void AddAssistantTurn(string message, string correlationId, ContentType contentType)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            throw new ArgumentException("Message cannot be empty", nameof(message));
        }

        var turn = new ConversationTurn
        {
            Role = "assistant",
            Content = message,
            CorrelationId = correlationId,
            ContentType = contentType,
            Timestamp = DateTimeOffset.UtcNow
        };

        _session.ConversationHistory.Add(turn);
        _session.ContextWindow.AddAssistantMessage(message);
        _session.UpdateActivity();

        // Evict oldest turns if context window exceeds limit
        EvictOldTurnsIfNeeded();
    }

    /// <summary>
    /// Adds a system message to full history only (NOT added to context window).
    /// Used for notifications like "Workbook changed to Budget.xlsx".
    /// </summary>
    public void AddSystemMessage(string message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            throw new ArgumentException("Message cannot be empty", nameof(message));
        }

        var turn = new ConversationTurn
        {
            Role = "system",
            Content = message,
            CorrelationId = string.Empty,
            ContentType = ContentType.SystemMessage,
            Timestamp = DateTimeOffset.UtcNow
        };

        _session.ConversationHistory.Add(turn);
        _session.UpdateActivity();
        // Note: System messages are NOT added to _session.ContextWindow
        // They're for UI display only
    }

    /// <summary>
    /// Gets the context window for LLM invocation (last N turns).
    /// Returns a Semantic Kernel ChatHistory instance.
    /// </summary>
    public ChatHistory GetContextForLLM()
    {
        return _session.ContextWindow;
    }

    /// <summary>
    /// Gets the complete conversation history for UI display.
    /// Includes all turns: user, assistant, and system messages.
    /// </summary>
    public List<ConversationTurn> GetFullHistory()
    {
        return new List<ConversationTurn>(_session.ConversationHistory);
    }

    /// <summary>
    /// Clears all conversation history and resets context window.
    /// </summary>
    public void Clear()
    {
        _session.ConversationHistory.Clear();
        _session.ContextWindow.Clear();
    }

    /// <summary>
    /// Evicts oldest user/assistant turn pairs when context window exceeds the configured limit.
    /// Removes 2 messages at a time (one user + one assistant) to maintain conversation coherence.
    /// </summary>
    private void EvictOldTurnsIfNeeded()
    {
        var maxTurns = _options.MaxContextTurns;
        
        // Each "turn" is a user message + assistant response = 2 messages
        var maxMessages = maxTurns * 2;

        while (_session.ContextWindow.Count > maxMessages)
        {
            // Remove the oldest 2 messages (1 user + 1 assistant)
            _session.ContextWindow.RemoveAt(0);
            if (_session.ContextWindow.Count > 0)
            {
                _session.ContextWindow.RemoveAt(0);
            }
        }
    }
}
