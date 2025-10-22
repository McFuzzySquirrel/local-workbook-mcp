namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Represents the structured response from the Semantic Kernel agent.
/// </summary>
public class AgentResponse
{
    /// <summary>
    /// Unique response identifier.
    /// </summary>
    public Guid ResponseId { get; set; } = Guid.NewGuid();

    /// <summary>
    /// Links to the query that triggered this response.
    /// </summary>
    public required string CorrelationId { get; set; }

    /// <summary>
    /// Primary response text.
    /// </summary>
    public required string Content { get; set; }

    /// <summary>
    /// Type of response (Text, Table, Error, Clarification).
    /// </summary>
    public ContentType ContentType { get; set; } = ContentType.Text;

    /// <summary>
    /// Structured table if ContentType is Table.
    /// </summary>
    public TableData? TableData { get; set; }

    /// <summary>
    /// Follow-up query suggestions.
    /// </summary>
    public List<string> Suggestions { get; set; } = new();

    /// <summary>
    /// MCP tools called by agent.
    /// </summary>
    public List<ToolInvocation> ToolsInvoked { get; set; } = new();

    /// <summary>
    /// Time taken to generate response (milliseconds).
    /// </summary>
    public int ProcessingTimeMs { get; set; }

    /// <summary>
    /// LLM model identifier.
    /// </summary>
    public string? ModelUsed { get; set; }

    /// <summary>
    /// Error info if response is error.
    /// </summary>
    public SanitizedError? Error { get; set; }
}
