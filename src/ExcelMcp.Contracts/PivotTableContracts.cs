namespace ExcelMcp.Contracts;

public sealed record PivotTableArguments(
    string Worksheet,
    string? PivotTable = null,
    bool IncludeFilters = true,
    int MaxRows = 100
);

public sealed record PivotTableResult(
    IReadOnlyList<PivotTableInfo> PivotTables
);

public sealed record PivotTableInfo(
    string Name,
    string WorksheetName,
    string SourceWorksheet,
    string SourceRange,
    IReadOnlyList<PivotFieldInfo> RowFields,
    IReadOnlyList<PivotFieldInfo> ColumnFields,
    IReadOnlyList<PivotFieldInfo> DataFields,
    IReadOnlyList<PivotFieldInfo> FilterFields,
    IReadOnlyList<PivotDataRow> Data
);

public sealed record PivotFieldInfo(
    string Name,
    string SourceName,
    string Function
);

public sealed record PivotDataRow(
    IReadOnlyDictionary<string, string?> Values
);
