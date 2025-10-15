namespace ExcelMcp.Contracts;

public sealed record ExcelSearchArguments(
    string Query,
    string? Worksheet = null,
    string? Table = null,
    int? Limit = null,
    bool CaseSensitive = false
);

public sealed record ExcelSearchResult(
    IReadOnlyList<ExcelRowResult> Rows,
    bool HasMore
);

public sealed record ExcelRowResult(
    string WorksheetName,
    string? TableName,
    int RowNumber,
    IReadOnlyDictionary<string, string?> Values
);
