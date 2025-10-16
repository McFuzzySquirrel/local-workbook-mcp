namespace ExcelMcp.ChatWeb.Options;

public sealed class ExcelMcpOptions
{
    public const string SectionName = "ExcelMcp";

    public string WorkbookPath { get; set; } = string.Empty;

    public string ServerPath { get; set; } = string.Empty;
}
