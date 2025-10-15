namespace ExcelMcp.Contracts;

public sealed record WorkbookMetadata(
    string WorkbookPath,
    IReadOnlyList<WorksheetMetadata> Worksheets,
    DateTimeOffset LastLoadedUtc
);

public sealed record WorksheetMetadata(
    string Name,
    IReadOnlyList<TableMetadata> Tables,
    IReadOnlyList<string> ColumnHeaders
);

public sealed record TableMetadata(
    string Name,
    string WorksheetName,
    IReadOnlyList<string> ColumnHeaders,
    int RowCount
);
