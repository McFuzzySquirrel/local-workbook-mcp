using ExcelMcp.Contracts;

namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Represents the currently loaded workbook and its metadata.
/// </summary>
public class WorkbookContext
{
    /// <summary>
    /// Absolute file path to the .xlsx file.
    /// </summary>
    public required string WorkbookPath { get; set; }

    /// <summary>
    /// Display name (filename without path).
    /// </summary>
    public required string WorkbookName { get; set; }

    /// <summary>
    /// When the workbook was loaded (UTC).
    /// </summary>
    public DateTimeOffset LoadedAt { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>
    /// Structured info from MCP server (sheets, tables, counts).
    /// </summary>
    public WorkbookMetadata? Metadata { get; set; }

    /// <summary>
    /// Whether workbook loaded successfully.
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// Sanitized error if load failed.
    /// </summary>
    public string? ErrorMessage { get; set; }
}
