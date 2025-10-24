namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Represents a single message exchange in the conversation (user query or agent response).
/// </summary>
public class ConversationTurn
{
    /// <summary>
    /// Unique identifier for the turn.
    /// </summary>
    public Guid Id { get; set; } = Guid.NewGuid();

    /// <summary>
    /// Role of the message sender: "user", "assistant", or "system".
    /// </summary>
    public required string Role { get; set; }

    /// <summary>
    /// The message text or structured data.
    /// </summary>
    public required string Content { get; set; }

    /// <summary>
    /// When the turn occurred (UTC).
    /// </summary>
    public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>
    /// Links related operations across logs.
    /// </summary>
    public required string CorrelationId { get; set; }

    /// <summary>
    /// Type of content (Text, Table, Error, SystemMessage, Clarification).
    /// </summary>
    public ContentType ContentType { get; set; } = ContentType.Text;

    /// <summary>
    /// Optional additional data.
    /// </summary>
    public Dictionary<string, object>? Metadata { get; set; }
}
