namespace ExcelMcp.ChatWeb.Options;

/// <summary>
/// Configuration options for Semantic Kernel.
/// </summary>
public class SemanticKernelOptions
{
    /// <summary>
    /// LLM model identifier (e.g., "phi-4-mini-reasoning").
    /// </summary>
    public string Model { get; set; } = "phi-4-mini-reasoning";

    /// <summary>
    /// Base URL for local LLM server (e.g., "http://localhost:1234").
    /// </summary>
    public string BaseUrl { get; set; } = "http://localhost:1234";

    /// <summary>
    /// API key (not needed for local models).
    /// </summary>
    public string ApiKey { get; set; } = "not-needed-for-local";

    /// <summary>
    /// Query timeout in seconds (default 30).
    /// </summary>
    public int TimeoutSeconds { get; set; } = 30;
}
