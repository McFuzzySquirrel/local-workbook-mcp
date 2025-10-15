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
