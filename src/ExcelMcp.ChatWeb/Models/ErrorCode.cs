namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Categorized error types for agent errors.
/// </summary>
public enum ErrorCode
{
    /// <summary>
    /// Could not open or read workbook.
    /// </summary>
    WorkbookLoadFailed,

    /// <summary>
    /// Query exceeded 30s limit.
    /// </summary>
    QueryTimeout,

    /// <summary>
    /// LLM server not responding.
    /// </summary>
    ModelUnresponsive,

    /// <summary>
    /// Query could not be parsed or understood.
    /// </summary>
    InvalidQuery,

    /// <summary>
    /// MCP tool invocation failed.
    /// </summary>
    McpToolError,

    /// <summary>
    /// Unexpected error.
    /// </summary>
    UnknownError
}
