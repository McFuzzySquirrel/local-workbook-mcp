using System.Collections.Generic;

namespace ExcelMcp.ChatWeb.Models;

public sealed record PreviewRequestDto
{
    public required string Worksheet { get; init; }
    public string? Table { get; init; }
    public int? Rows { get; init; }
    public string? Cursor { get; init; }
}

public sealed record PreviewRowDto(int RowNumber, IReadOnlyList<string?> Values);

public sealed record PreviewResponseDto(
    bool Success,
    string? Error,
    string? Worksheet,
    string? Table,
    int Offset,
    IReadOnlyList<string> Headers,
    IReadOnlyList<PreviewRowDto> Rows,
    bool HasMore,
    string? NextCursor,
    string? Csv
);

internal sealed record PreviewToolPayload(
    string Worksheet,
    string? Table,
    int Offset,
    bool HasMore,
    string? NextCursor,
    IReadOnlyList<string> Headers,
    IReadOnlyList<PreviewToolRowPayload> Rows
);

internal sealed record PreviewToolRowPayload(int RowNumber, IReadOnlyList<string?> Values);
