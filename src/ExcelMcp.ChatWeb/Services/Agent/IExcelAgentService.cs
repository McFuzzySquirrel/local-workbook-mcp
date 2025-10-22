using ExcelMcp.ChatWeb.Models;

namespace ExcelMcp.ChatWeb.Services.Agent;

/// <summary>
/// Main service interface for agent interactions with Excel workbooks.
/// </summary>
public interface IExcelAgentService
{
    /// <summary>
    /// Processes a user query and returns an agent response.
    /// </summary>
    /// <param name="query">User's natural language question.</param>
    /// <param name="session">Current workbook session with context.</param>
    /// <param name="cancellationToken">For timeout/cancellation support.</param>
    /// <returns>Agent response with content, table data, or error.</returns>
    Task<AgentResponse> ProcessQueryAsync(
        string query,
        WorkbookSession session,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Loads a new workbook and updates session context.
    /// </summary>
    /// <param name="filePath">Absolute path to .xlsx file.</param>
    /// <param name="session">Current session to update.</param>
    /// <param name="cancellationToken">For cancellation support.</param>
    /// <returns>Workbook context with metadata or error info.</returns>
    Task<WorkbookContext> LoadWorkbookAsync(
        string filePath,
        WorkbookSession session,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// Clears conversation history and starts fresh session.
    /// </summary>
    /// <param name="session">Session to clear.</param>
    Task ClearConversationAsync(WorkbookSession session);

    /// <summary>
    /// Generates relevant follow-up questions based on current context.
    /// </summary>
    /// <param name="session">Current session with workbook and conversation.</param>
    /// <param name="maxSuggestions">Number of suggestions to return.</param>
    /// <returns>List of suggested query strings.</returns>
    Task<List<string>> GetSuggestedQueriesAsync(
        WorkbookSession session,
        int maxSuggestions = 3);
}
