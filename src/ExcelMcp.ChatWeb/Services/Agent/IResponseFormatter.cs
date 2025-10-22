using ExcelMcp.ChatWeb.Models;

namespace ExcelMcp.ChatWeb.Services.Agent;

/// <summary>
/// Formats agent responses for UI display.
/// </summary>
public interface IResponseFormatter
{
    /// <summary>
    /// Converts table data to HTML table string.
    /// </summary>
    /// <param name="tableData">Structured table with columns and rows.</param>
    /// <returns>HTML string with styled table.</returns>
    string FormatAsHtmlTable(TableData tableData);

    /// <summary>
    /// Formats a text response with markdown-like styling.
    /// </summary>
    /// <param name="content">Plain text or simple markdown.</param>
    /// <returns>Formatted HTML string.</returns>
    string FormatAsText(string content);

    /// <summary>
    /// Sanitizes error for user display (removes sensitive data).
    /// </summary>
    /// <param name="exception">The actual exception.</param>
    /// <param name="correlationId">Tracking ID.</param>
    /// <returns>Sanitized error with generic message.</returns>
    SanitizedError SanitizeErrorMessage(Exception exception, string correlationId);
}
