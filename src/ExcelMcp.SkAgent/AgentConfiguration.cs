namespace ExcelMcp.SkAgent;

public sealed class AgentConfiguration
{
    public required string BaseUrl { get; init; }
    public required string ModelId { get; init; }
    public required string ApiKey { get; init; }
}
