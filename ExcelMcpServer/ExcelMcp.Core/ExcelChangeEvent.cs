using System;

namespace ExcelMcp.Core
{
    public record ExcelChangeEvent(
        string Sheet,
        string Table,
        string ChangeType,
        int? Row,
        string? Column,
        string? OldValue,
        string? NewValue,
        DateTime Timestamp);
}