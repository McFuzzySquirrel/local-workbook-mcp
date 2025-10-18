namespace ExcelMcp.ChatWeb.Options;

public sealed class LlmStudioOptions
{
    public const string SectionName = "LlmStudio";

    public string BaseUrl { get; set; } = "http://localhost:1234";

    public string Model { get; set; } = "phi-4-mini-reasoning";

    public double Temperature { get; set; } = 0.2;
}
