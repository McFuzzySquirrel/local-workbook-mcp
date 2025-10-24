using Microsoft.SemanticKernel.ChatCompletion;

namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Represents an active user session with conversation history and workbook state.
/// </summary>
public class WorkbookSession
{
    /// <summary>
    /// Unique session identifier.
    /// </summary>
    public Guid SessionId { get; set; } = Guid.NewGuid();

    /// <summary>
    /// Session creation time (UTC).
    /// </summary>
    public DateTimeOffset StartedAt { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>
    /// Most recent interaction time (UTC).
    /// </summary>
    public DateTimeOffset LastActivityAt { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>
    /// Active workbook, null if none loaded.
    /// </summary>
    public WorkbookContext? CurrentContext { get; set; }

    /// <summary>
    /// Full conversation (all turns for UI display).
    /// </summary>
    public List<ConversationTurn> ConversationHistory { get; set; } = new();

    /// <summary>
    /// SK ChatHistory with last 20 turns for LLM context.
    /// </summary>
    public ChatHistory ContextWindow { get; set; } = new();

    /// <summary>
    /// History of loaded workbooks (for multi-workbook sessions).
    /// </summary>
    public List<WorkbookContext> PreviousContexts { get; set; } = new();

    /// <summary>
    /// Update last activity timestamp.
    /// </summary>
    public void UpdateActivity()
    {
        LastActivityAt = DateTimeOffset.UtcNow;
    }

    /// <summary>
    /// Load new workbook and preserve history.
    /// </summary>
    public void LoadNewWorkbook(WorkbookContext context)
    {
        if (CurrentContext != null)
        {
            PreviousContexts.Add(CurrentContext);
        }
        CurrentContext = context;
        UpdateActivity();
    }
}
