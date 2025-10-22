namespace ExcelMcp.ChatWeb.Models;

/// <summary>
/// Structured table data for display in chat interface.
/// </summary>
public class TableData
{
    /// <summary>
    /// Column headers.
    /// </summary>
    public List<string> Columns { get; set; } = new();

    /// <summary>
    /// Data rows (each row is list of cell values).
    /// </summary>
    public List<List<string>> Rows { get; set; } = new();

    /// <summary>
    /// Table metadata (sheet name, row count, etc.).
    /// </summary>
    public TableMetadata? Metadata { get; set; }
}

/// <summary>
/// Metadata about the table data.
/// </summary>
public class TableMetadata
{
    /// <summary>
    /// Source sheet name.
    /// </summary>
    public string? SheetName { get; set; }

    /// <summary>
    /// Total row count in source (may be more than displayed).
    /// </summary>
    public int? RowCount { get; set; }

    /// <summary>
    /// Whether data was truncated for display.
    /// </summary>
    public bool IsTruncated { get; set; }
}
