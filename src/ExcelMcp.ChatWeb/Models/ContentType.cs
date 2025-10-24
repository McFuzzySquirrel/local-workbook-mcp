namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Type of content in a conversation turn or agent response.
/// </summary>
public enum ContentType
{
    /// <summary>
    /// Plain text response.
    /// </summary>
    Text,

    /// <summary>
    /// Structured data table.
    /// </summary>
    Table,

    /// <summary>
    /// Error message.
    /// </summary>
    Error,

    /// <summary>
    /// System notification (e.g., workbook switch marker).
    /// </summary>
    SystemMessage,

    /// <summary>
    /// Agent asking for user clarification.
    /// </summary>
    Clarification
}
