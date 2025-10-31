using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using ExcelMcp.Server.Excel;

#pragma warning disable SKEXP0010

namespace ExcelMcp.SkAgent;

public sealed class ExcelAgent
{
    private readonly string _workbookPath;
    private readonly AgentConfiguration _config;
    private Kernel? _kernel;
    private IChatCompletionService? _chatService;
    private ExcelWorkbookService? _workbookService;
    
    public List<string> DebugLog { get; } = new();

    public ExcelAgent(string workbookPath, AgentConfiguration config)
    {
        _workbookPath = workbookPath;
        _config = config;
    }

    public Task InitializeAsync(CancellationToken cancellationToken = default)
    {
        _workbookService = new ExcelWorkbookService(_workbookPath);

        var builder = Kernel.CreateBuilder();
        
        // Configure HttpClient for local LLM compatibility
        var httpClient = new HttpClient();
        httpClient.BaseAddress = new Uri(_config.BaseUrl);
        httpClient.Timeout = TimeSpan.FromMinutes(5);
        
        builder.AddOpenAIChatCompletion(
            modelId: _config.ModelId,
            apiKey: _config.ApiKey,
            httpClient: httpClient);

        // Add function call logging
        builder.Plugins.AddFromObject(new ExcelPlugin(_workbookService, DebugLog), "Excel");

        _kernel = builder.Build();
        _chatService = _kernel.GetRequiredService<IChatCompletionService>();

        return Task.CompletedTask;
    }

    public async Task ProcessMessageAsync(ChatHistory history, CancellationToken cancellationToken = default)
    {
        if (_kernel is null || _chatService is null)
        {
            throw new InvalidOperationException("Agent not initialized. Call InitializeAsync first.");
        }

        DebugLog.Clear(); // Clear previous debug logs

        var executionSettings = new OpenAIPromptExecutionSettings
        {
            ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions,
            Temperature = 0.1,  // Lower temperature for more accurate, less creative responses
            MaxTokens = 2000,
            TopP = 0.1  // Focus on most likely tokens
        };

        try
        {
            DebugLog.Add("🔄 Sending request to LLM...");
            
            var response = await _chatService.GetChatMessageContentAsync(
                history,
                executionSettings,
                _kernel,
                cancellationToken);

            if (DebugLog.Count == 1)
            {
                DebugLog.Add("⚠️  No tools were called - LLM answered directly");
            }

            history.Add(response);
        }
        catch (Exception ex)
        {
            var errorMessage = $"Error communicating with LLM: {ex.Message}";
            DebugLog.Add($"❌ Error: {ex.Message}");
            history.AddAssistantMessage(errorMessage);
            throw new InvalidOperationException(errorMessage, ex);
        }
    }
}
