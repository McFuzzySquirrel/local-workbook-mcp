namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Represents a call to an MCP tool via Semantic Kernel plugin.
/// </summary>
public class ToolInvocation
{
    /// <summary>
    /// MCP tool identifier (e.g., "excel-list-structure").
    /// </summary>
    public required string ToolName { get; set; }

    /// <summary>
    /// SK plugin method name.
    /// </summary>
    public required string PluginMethod { get; set; }

    /// <summary>
    /// Parameters passed to tool.
    /// </summary>
    public Dictionary<string, object> InputParameters { get; set; } = new();

    /// <summary>
    /// Brief summary of result (not full data).
    /// </summary>
    public string? OutputSummary { get; set; }

    /// <summary>
    /// Execution timestamp (UTC).
    /// </summary>
    public DateTimeOffset InvokedAt { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>
    /// Execution time in milliseconds.
    /// </summary>
    public int DurationMs { get; set; }

    /// <summary>
    /// Whether tool call succeeded.
    /// </summary>
    public bool Success { get; set; }

    /// <summary>
    /// Sanitized error if failed.
    /// </summary>
    public string? ErrorMessage { get; set; }
}
