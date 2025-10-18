using System.Collections.Generic;

namespace ExcelMcp.Contracts;

public sealed record ExcelResourceDescriptor(
    Uri Uri,
    string Name,
    string? Description,
    string MimeType
);

public sealed record ExcelResourceContent(
    Uri Uri,
    string MimeType,
    string Text
);

public sealed record ExcelPreviewArguments(
    string Worksheet,
    string? Table = null,
    int? Rows = null,
    string? Cursor = null
);

public sealed record ExcelPreviewRow(
    int RowNumber,
    IReadOnlyList<string?> Values
);

public sealed record ExcelPreviewResult(
    string Worksheet,
    string? Table,
    IReadOnlyList<string> Headers,
    IReadOnlyList<ExcelPreviewRow> Rows,
    int Offset,
    bool HasMore,
    string? NextCursor,
    string Csv
);
