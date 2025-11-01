using ExcelMcp.ChatWeb.Components;
using ExcelMcp.ChatWeb.Logging;
using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.ChatWeb.Services;
using ExcelMcp.ChatWeb.Services.Agent;
using ExcelMcp.ChatWeb.Services.Plugins;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using Microsoft.SemanticKernel;
using Serilog;

// Configure Serilog from appsettings.json
Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json")
        .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT") ?? "Production"}.json", optional: true)
        .Build())
    .CreateLogger();

try
{
    Log.Information("Starting ExcelMcp.ChatWeb application");

    var builder = WebApplication.CreateBuilder(args);

    // Use Serilog for logging
    builder.Host.UseSerilog();

EnsureExcelMcpConfiguration(builder);

builder.Services.AddOptions<LlmStudioOptions>()
    .Bind(builder.Configuration.GetSection(LlmStudioOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddOptions<ExcelMcpOptions>()
    .Bind(builder.Configuration.GetSection(ExcelMcpOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddOptions<SemanticKernelOptions>()
    .Bind(builder.Configuration.GetSection("SemanticKernel"))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddOptions<ConversationOptions>()
    .Bind(builder.Configuration.GetSection("Conversation"))
    .ValidateDataAnnotations()
    .ValidateOnStart();

// Auto-detect running model from LM Studio (like CLI does)
var detectedModel = await DetectRunningModelAsync(builder.Configuration.GetSection("SemanticKernel")["BaseUrl"] ?? "http://localhost:1234/v1");
if (!string.IsNullOrEmpty(detectedModel) && detectedModel != "unknown")
{
    Log.Information("Auto-detected running model: {ModelName}", detectedModel);
    // Override the configured model with the detected one
    builder.Services.Configure<SemanticKernelOptions>(options => options.Model = detectedModel);
}
else
{
    var configuredModel = builder.Configuration.GetSection("SemanticKernel")["Model"];
    Log.Warning("Could not detect running model, using configured model: {ModelName}", configuredModel);
}

// Register plugins first
builder.Services.AddSingleton<WorkbookStructurePlugin>();
builder.Services.AddSingleton<WorkbookSearchPlugin>();
builder.Services.AddSingleton<DataRetrievalPlugin>();

// Semantic Kernel with OpenAI chat completion
builder.Services.AddSingleton(serviceProvider =>
{
    var skOptions = serviceProvider.GetRequiredService<IOptions<SemanticKernelOptions>>().Value;
    
    var kernelBuilder = Kernel.CreateBuilder();
    
    // Add OpenAI chat completion for local LLM (LM Studio, Ollama, etc.)
    // Use HttpClient to bypass OpenAI SDK's strict response validation
    var httpClient = new HttpClient();
    kernelBuilder.AddOpenAIChatCompletion(
        modelId: skOptions.Model,
        apiKey: skOptions.ApiKey,
        endpoint: new Uri(skOptions.BaseUrl),
        httpClient: httpClient);
    
    // Add plugins
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<WorkbookStructurePlugin>());
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<WorkbookSearchPlugin>());
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<DataRetrievalPlugin>());
    
    return kernelBuilder.Build();
});

// Agent services
builder.Services.AddSingleton<AgentLogger>();
builder.Services.AddScoped<IExcelAgentService, ExcelAgentService>();
builder.Services.AddScoped<IConversationManager, ConversationManager>();
builder.Services.AddSingleton<IResponseFormatter, ResponseFormatter>();

// Session state (per Blazor circuit)
builder.Services.AddScoped<WorkbookSession>();

builder.Services.AddSingleton<McpClientHost>();
builder.Services.AddSingleton<IMcpClient>(static sp => sp.GetRequiredService<McpClientHost>());
builder.Services.AddSingleton<IHostedService>(static sp => sp.GetRequiredService<McpClientHost>());

builder.Services.AddHttpClient<ILlmStudioClient, LlmStudioClient>((serviceProvider, client) =>
{
    var options = serviceProvider.GetRequiredService<IOptions<LlmStudioOptions>>().Value;
    client.BaseAddress = new Uri(options.BaseUrl, UriKind.Absolute);
    client.Timeout = TimeSpan.FromSeconds(120);
});

builder.Services.AddSingleton<ChatService>();

builder.Services.AddSingleton(static sp =>
{
    var options = sp.GetRequiredService<IOptions<LlmStudioOptions>>().Value;
    return new ModelInfoDto(options.Model, options.BaseUrl);
});

// Add Blazor Server services
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

    var app = builder.Build();

    // Add correlation ID middleware for request tracking
    app.UseCorrelationId();

    // Add Serilog request logging
    app.UseSerilogRequestLogging();

    app.UseStaticFiles();
    app.UseAntiforgery();

    // Legacy API endpoints (keep for backwards compatibility)
    app.MapPost("/api/chat", async (ChatRequestDto request, ChatService chatService, CancellationToken cancellationToken) =>
    {
        var response = await chatService.HandleAsync(request, cancellationToken).ConfigureAwait(false);
        return Results.Json(response);
    });

    app.MapGet("/health", () => Results.Ok(new { status = "ok" }));

    app.MapGet("/api/model", (ModelInfoDto info) => Results.Json(info));

    // Map Blazor components
    app.MapRazorComponents<App>()
        .AddInteractiveServerRenderMode();

    Log.Information("ExcelMcp.ChatWeb application started successfully");
    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
    throw;
}
finally
{
    Log.CloseAndFlush();
}

// Auto-detect running model from LM Studio API
static async Task<string> DetectRunningModelAsync(string baseUrl)
{
    try
    {
        using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(5) };
        var modelsUrl = baseUrl.Replace("/v1", "") + "/v1/models";
        
        var response = await httpClient.GetStringAsync(modelsUrl);
        
        // Use proper JSON parsing
        var jsonDoc = System.Text.Json.JsonDocument.Parse(response);
        if (jsonDoc.RootElement.TryGetProperty("data", out var dataArray))
        {
            if (dataArray.GetArrayLength() > 0)
            {
                var firstModel = dataArray[0];
                if (firstModel.TryGetProperty("id", out var idElement))
                {
                    return idElement.GetString() ?? "unknown";
                }
            }
        }
        
        return "unknown";
    }
    catch (Exception ex)
    {
        Log.Warning(ex, "Failed to auto-detect model from {BaseUrl}", baseUrl);
        return "unknown";
    }
}

static void EnsureExcelMcpConfiguration(WebApplicationBuilder builder)
{
    var serverConfigKey = $"{ExcelMcpOptions.SectionName}:ServerPath";
    var serverPath = builder.Configuration[serverConfigKey];

    if (string.IsNullOrWhiteSpace(serverPath))
    {
        serverPath = Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER");
    }

    // Only prompt for server path if not configured (workbook is selected via UI now)
    var interactive = !Console.IsInputRedirected && !Console.IsOutputRedirected && !Console.IsErrorRedirected;

    if (string.IsNullOrWhiteSpace(serverPath) && interactive)
    {
        var detectedServer = TryFindServerExecutable();

        if (!string.IsNullOrWhiteSpace(detectedServer))
        {
            Console.Write($"Enter the MCP server executable path [{detectedServer}]: ");
            var input = Console.ReadLine();
            serverPath = string.IsNullOrWhiteSpace(input) ? detectedServer : input.Trim();
        }
        else
        {
            Console.Write("Enter the MCP server executable path: ");
            serverPath = Console.ReadLine()?.Trim();
        }
    }

    if (!string.IsNullOrWhiteSpace(serverPath))
    {
        var resolved = Path.GetFullPath(serverPath);
        var overrides = new Dictionary<string, string?>
        {
            [serverConfigKey] = resolved
        };
        Environment.SetEnvironmentVariable("EXCEL_MCP_SERVER", resolved);
        builder.Configuration.AddInMemoryCollection(overrides);
    }
}

static string? TryFindServerExecutable()
{
    var baseDirectory = AppContext.BaseDirectory;
    var fileName = OperatingSystem.IsWindows() ? "ExcelMcp.Server.exe" : "ExcelMcp.Server";

    var candidates = new[]
    {
        Path.Combine(baseDirectory, fileName),
        Path.Combine(baseDirectory, "ExcelMcp.Server", fileName),
        Path.Combine(baseDirectory, "..", "ExcelMcp.Server", fileName),
        Path.Combine(baseDirectory, "..", "..", "ExcelMcp.Server", fileName)
    };

    foreach (var candidate in candidates.Select(Path.GetFullPath))
    {
        if (File.Exists(candidate))
        {
            return candidate;
        }
    }

    return null;
}
