namespace ExcelMcp.ChatWeb.Options;

public sealed class OllamaOptions
{
    public const string SectionName = "Ollama";

    public string BaseUrl { get; set; } = "http://localhost:11434";

    public string Model { get; set; } = "gemma3:1b";

    public double Temperature { get; set; } = 0.2;
}
