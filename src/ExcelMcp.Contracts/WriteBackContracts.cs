namespace ExcelMcp.Contracts;

/// <summary>
/// Arguments for updating a cell value in the workbook.
/// </summary>
public sealed record UpdateCellArguments(
    string Worksheet,
    string CellAddress,
    string Value,
    string? Reason = null
);

/// <summary>
/// Result of an update cell operation.
/// </summary>
public sealed record UpdateCellResult(
    string Worksheet,
    string CellAddress,
    string? PreviousValue,
    string NewValue,
    DateTimeOffset Timestamp,
    string? AuditId
);

/// <summary>
/// Arguments for adding a new worksheet to the workbook.
/// </summary>
public sealed record AddWorksheetArguments(
    string Name,
    int? Position = null,
    string? Reason = null
);

/// <summary>
/// Result of adding a new worksheet.
/// </summary>
public sealed record AddWorksheetResult(
    string Name,
    int Position,
    DateTimeOffset Timestamp,
    string? AuditId
);

/// <summary>
/// Arguments for adding an annotation (comment) to a cell.
/// </summary>
public sealed record AddAnnotationArguments(
    string Worksheet,
    string CellAddress,
    string Text,
    string? Author = null
);

/// <summary>
/// Result of adding an annotation.
/// </summary>
public sealed record AddAnnotationResult(
    string Worksheet,
    string CellAddress,
    string Text,
    string? Author,
    DateTimeOffset Timestamp,
    string? AuditId
);

/// <summary>
/// An entry in the audit trail tracking changes made to the workbook.
/// </summary>
public sealed record AuditEntry(
    string Id,
    string OperationType,
    string Description,
    DateTimeOffset Timestamp,
    string? Reason,
    IReadOnlyDictionary<string, string?> Details
);

/// <summary>
/// Arguments for querying the audit trail.
/// </summary>
public sealed record GetAuditTrailArguments(
    DateTimeOffset? Since = null,
    DateTimeOffset? Until = null,
    string? OperationType = null,
    int? Limit = null
);

/// <summary>
/// Result of querying the audit trail.
/// </summary>
public sealed record GetAuditTrailResult(
    IReadOnlyList<AuditEntry> Entries,
    int TotalCount
);
