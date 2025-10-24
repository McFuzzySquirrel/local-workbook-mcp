using System.Text;
using ExcelMcp.ChatWeb.Models;

namespace ExcelMcp.ChatWeb.Services.Agent;

/// <summary>
/// Formats agent responses for display in the UI.
/// Converts TableData to HTML tables, formats text content, and sanitizes errors.
/// </summary>
public class ResponseFormatter : IResponseFormatter
{
    /// <summary>
    /// Formats a TableData object as an HTML table with styling.
    /// </summary>
    public string FormatAsHtmlTable(TableData tableData)
    {
        if (tableData == null)
        {
            throw new ArgumentNullException(nameof(tableData));
        }

        var html = new StringBuilder();
        
        html.AppendLine("<div class=\"table-container\">");
        
        // Add table metadata if present
        if (tableData.Metadata != null)
        {
            html.AppendLine($"<div class=\"table-header\">");
            if (!string.IsNullOrEmpty(tableData.Metadata.SheetName))
            {
                html.AppendLine($"  <span class=\"sheet-name\">Sheet: {EscapeHtml(tableData.Metadata.SheetName)}</span>");
            }
            html.AppendLine($"</div>");
        }

        html.AppendLine("<table class=\"data-table\">");
        
        // Table header
        if (tableData.Columns?.Any() == true)
        {
            html.AppendLine("  <thead>");
            html.AppendLine("    <tr>");
            foreach (var column in tableData.Columns)
            {
                html.AppendLine($"      <th>{EscapeHtml(column)}</th>");
            }
            html.AppendLine("    </tr>");
            html.AppendLine("  </thead>");
        }

        // Table body
        html.AppendLine("  <tbody>");
        if (tableData.Rows?.Any() == true)
        {
            foreach (var row in tableData.Rows)
            {
                html.AppendLine("    <tr>");
                foreach (var cell in row)
                {
                    var cellValue = cell?.ToString() ?? string.Empty;
                    html.AppendLine($"      <td>{EscapeHtml(cellValue)}</td>");
                }
                html.AppendLine("    </tr>");
            }
        }
        else
        {
            html.AppendLine("    <tr>");
            html.AppendLine($"      <td colspan=\"{tableData.Columns?.Count ?? 1}\" class=\"empty-state\">No data available</td>");
            html.AppendLine("    </tr>");
        }
        html.AppendLine("  </tbody>");
        
        html.AppendLine("</table>");
        
        // Add row count footer
        var rowCount = tableData.Rows?.Count ?? 0;
        html.AppendLine($"<div class=\"table-footer\">");
        html.AppendLine($"  <span class=\"row-count\">{rowCount} row{(rowCount != 1 ? "s" : "")} displayed</span>");
        html.AppendLine($"</div>");
        
        html.AppendLine("</div>");

        return html.ToString();
    }

    /// <summary>
    /// Formats text content, converting markdown-like syntax to HTML.
    /// Supports: bold, italic, code blocks, inline code, lists.
    /// </summary>
    public string FormatAsText(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return string.Empty;
        }

        var html = new StringBuilder();
        var lines = content.Split('\n');
        bool inCodeBlock = false;
        bool inList = false;

        foreach (var line in lines)
        {
            var trimmedLine = line.TrimStart();

            // Code blocks (```language or ```)
            if (trimmedLine.StartsWith("```"))
            {
                if (inCodeBlock)
                {
                    html.AppendLine("</code></pre>");
                    inCodeBlock = false;
                }
                else
                {
                    if (inList)
                    {
                        html.AppendLine("</ul>");
                        inList = false;
                    }
                    var language = trimmedLine.Length > 3 ? trimmedLine.Substring(3).Trim() : "";
                    html.AppendLine($"<pre><code class=\"language-{EscapeHtml(language)}\">");
                    inCodeBlock = true;
                }
                continue;
            }

            if (inCodeBlock)
            {
                html.AppendLine(EscapeHtml(line));
                continue;
            }

            // Lists (- item or * item)
            if (trimmedLine.StartsWith("- ") || trimmedLine.StartsWith("* "))
            {
                if (!inList)
                {
                    html.AppendLine("<ul>");
                    inList = true;
                }
                var listItem = trimmedLine.Substring(2);
                html.AppendLine($"  <li>{FormatInlineMarkdown(listItem)}</li>");
                continue;
            }
            else if (inList)
            {
                html.AppendLine("</ul>");
                inList = false;
            }

            // Empty lines create paragraphs
            if (string.IsNullOrWhiteSpace(trimmedLine))
            {
                if (html.Length > 0 && !html.ToString().EndsWith("<br>\n"))
                {
                    html.AppendLine("<br>");
                }
                continue;
            }

            // Regular text with inline markdown
            html.AppendLine($"<p>{FormatInlineMarkdown(trimmedLine)}</p>");
        }

        // Close any open tags
        if (inCodeBlock)
        {
            html.AppendLine("</code></pre>");
        }
        if (inList)
        {
            html.AppendLine("</ul>");
        }

        return html.ToString();
    }

    /// <summary>
    /// Sanitizes an exception into a user-friendly error message.
    /// Maps exception types to appropriate error codes and messages.
    /// </summary>
    public SanitizedError SanitizeErrorMessage(Exception exception, string correlationId)
    {
        if (exception == null)
        {
            throw new ArgumentNullException(nameof(exception));
        }

        // Map exception types to error codes and user-friendly messages
        var (errorCode, userMessage, canRetry, suggestedAction) = exception switch
        {
            FileNotFoundException => 
                (ErrorCode.WorkbookLoadFailed, 
                 "The specified workbook file could not be found.", 
                 false,
                 "Please verify the file path and try again."),

            UnauthorizedAccessException => 
                (ErrorCode.WorkbookLoadFailed, 
                 "Permission denied while accessing the workbook file.", 
                 false,
                 "Check file permissions and ensure the file is not open in another application."),

            TimeoutException => 
                (ErrorCode.QueryTimeout, 
                 "The query took too long to complete.", 
                 true,
                 "Try a more specific query or retry the operation."),

            OperationCanceledException => 
                (ErrorCode.QueryTimeout, 
                 "The operation was cancelled or timed out.", 
                 true,
                 "Please try again with a simpler query."),

            InvalidOperationException when exception.Message.Contains("workbook") => 
                (ErrorCode.InvalidQuery, 
                 "No workbook is currently loaded.", 
                 false,
                 "Please load a workbook before asking questions."),

            ArgumentException => 
                (ErrorCode.InvalidQuery, 
                 "The query contains invalid parameters.", 
                 false,
                 "Please rephrase your question and try again."),

            HttpRequestException => 
                (ErrorCode.ModelUnresponsive, 
                 "Unable to communicate with the AI model.", 
                 true,
                 "Ensure LM Studio or your local LLM server is running and try again."),

            _ => 
                (ErrorCode.UnknownError, 
                 "An unexpected error occurred while processing your request.", 
                 true,
                 "Please try again or contact support if the problem persists.")
        };

        return new SanitizedError
        {
            Message = userMessage,
            ErrorCode = errorCode,
            CorrelationId = correlationId,
            CanRetry = canRetry,
            SuggestedAction = suggestedAction,
            Timestamp = DateTimeOffset.UtcNow
        };
    }

    /// <summary>
    /// Formats inline markdown (bold, italic, inline code).
    /// </summary>
    private string FormatInlineMarkdown(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        // Escape HTML first
        text = EscapeHtml(text);

        // Inline code (`code`)
        text = System.Text.RegularExpressions.Regex.Replace(
            text, 
            @"`([^`]+)`", 
            "<code>$1</code>");

        // Bold (**text** or __text__)
        text = System.Text.RegularExpressions.Regex.Replace(
            text, 
            @"\*\*([^\*]+)\*\*", 
            "<strong>$1</strong>");
        text = System.Text.RegularExpressions.Regex.Replace(
            text, 
            @"__([^_]+)__", 
            "<strong>$1</strong>");

        // Italic (*text* or _text_)
        text = System.Text.RegularExpressions.Regex.Replace(
            text, 
            @"\*([^\*]+)\*", 
            "<em>$1</em>");
        text = System.Text.RegularExpressions.Regex.Replace(
            text, 
            @"_([^_]+)_", 
            "<em>$1</em>");

        return text;
    }

    /// <summary>
    /// Escapes HTML special characters.
    /// </summary>
    private string EscapeHtml(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }
}
