namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Represents a user-safe error message with troubleshooting reference.
/// </summary>
public class SanitizedError
{
    /// <summary>
    /// Generic user-friendly error message (no sensitive data).
    /// </summary>
    public required string Message { get; set; }

    /// <summary>
    /// Links to detailed log entry for troubleshooting.
    /// </summary>
    public required string CorrelationId { get; set; }

    /// <summary>
    /// When error occurred (UTC).
    /// </summary>
    public DateTimeOffset Timestamp { get; set; } = DateTimeOffset.UtcNow;

    /// <summary>
    /// Categorized error type.
    /// </summary>
    public ErrorCode ErrorCode { get; set; }

    /// <summary>
    /// Whether user can retry the operation.
    /// </summary>
    public bool CanRetry { get; set; }

    /// <summary>
    /// Helpful guidance for user.
    /// </summary>
    public string? SuggestedAction { get; set; }
}
