namespace ExcelMcp.Contracts;

/// <summary>Request to write a single cell value.</summary>
public sealed record WriteCellRequest(
    /// <summary>Path to the Excel workbook file.</summary>
    string WorkbookPath,
    /// <summary>Target worksheet name.</summary>
    string Worksheet,
    /// <summary>Cell address in A1 notation (e.g. "B4").</summary>
    string CellAddress,
    /// <summary>Value to write. Pass null to clear the cell.</summary>
    string? Value
);

/// <summary>A single cell+value pair within a WriteRangeRequest.</summary>
public sealed record CellUpdate(
    /// <summary>Cell address in A1 notation (e.g. "C12").</summary>
    string CellAddress,
    /// <summary>Value to write. Pass null to clear the cell.</summary>
    string? Value
);

/// <summary>Request to write multiple cells in one operation.</summary>
public sealed record WriteRangeRequest(
    /// <summary>Path to the Excel workbook file.</summary>
    string WorkbookPath,
    /// <summary>Target worksheet name.</summary>
    string Worksheet,
    /// <summary>Cells to update.</summary>
    IReadOnlyList<CellUpdate> Updates
);

/// <summary>Request to create a new worksheet.</summary>
public sealed record CreateWorksheetRequest(
    /// <summary>Path to the Excel workbook file.</summary>
    string WorkbookPath,
    /// <summary>Name for the new worksheet.</summary>
    string WorksheetName
);

/// <summary>Result returned by all write operations.</summary>
public sealed record WriteResult(
    /// <summary>True if the operation succeeded.</summary>
    bool Success,
    /// <summary>Human-readable summary of what was done.</summary>
    string Message,
    /// <summary>Path to the backup file created before writing.</summary>
    string? BackupPath = null
);
