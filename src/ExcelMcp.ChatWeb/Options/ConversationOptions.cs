namespace ExcelMcp.ChatWeb.Options;

/// <summary>
/// Configuration options for conversation management.
/// </summary>
public class ConversationOptions
{
    /// <summary>
    /// Maximum number of conversation turns to keep in context window (default 20).
    /// </summary>
    public int MaxContextTurns { get; set; } = 20;

    /// <summary>
    /// Maximum length of agent response in characters (default 10000).
    /// </summary>
    public int MaxResponseLength { get; set; } = 10000;

    /// <summary>
    /// Number of suggested queries to generate (default 3).
    /// </summary>
    public int SuggestedQueriesCount { get; set; } = 3;
}
