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

// Auto-detect running LLM provider: tries Ollama first, then LM Studio
var detected = await DetectProviderAsync();
if (detected is { } provider)
{
    Log.Information("Auto-detected LLM provider: model={Model} url={BaseUrl}", provider.Model, provider.BaseUrl);
    builder.Services.Configure<SemanticKernelOptions>(options =>
    {
        options.Model = provider.Model;
        options.BaseUrl = provider.BaseUrl;
    });
    builder.Services.Configure<LlmStudioOptions>(options =>
    {
        options.Model = provider.Model;
        options.BaseUrl = provider.BaseUrl.Replace("/v1", "");
    });
}
else
{
    var configuredModel = builder.Configuration.GetSection("SemanticKernel")["Model"];
    Log.Warning("No local LLM detected, using configured defaults: {ModelName}", configuredModel);
}

// Register plugins first
builder.Services.AddSingleton<WorkbookStructurePlugin>();
builder.Services.AddSingleton<WorkbookSearchPlugin>();
builder.Services.AddSingleton<DataRetrievalPlugin>();
builder.Services.AddSingleton<WorkbookWritePlugin>();

// Semantic Kernel with OpenAI chat completion
builder.Services.AddSingleton(serviceProvider =>
{
    var skOptions = serviceProvider.GetRequiredService<IOptions<SemanticKernelOptions>>().Value;
    
    var kernelBuilder = Kernel.CreateBuilder();
    
    // Add OpenAI chat completion for local LLM (LM Studio, Ollama, etc.)
    // Use HttpClient with extended timeout for slower models on Raspberry Pi
    var httpClient = new HttpClient
    {
        Timeout = TimeSpan.FromSeconds(skOptions.TimeoutSeconds) // Use configured timeout (480s)
    };
    
    kernelBuilder.AddOpenAIChatCompletion(
        modelId: skOptions.Model,
        apiKey: skOptions.ApiKey,
        endpoint: new Uri(skOptions.BaseUrl),
        httpClient: httpClient);
    
    // Add plugins
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<WorkbookStructurePlugin>());
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<WorkbookSearchPlugin>());
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<DataRetrievalPlugin>());
    kernelBuilder.Plugins.AddFromObject(serviceProvider.GetRequiredService<WorkbookWritePlugin>());
    
    return kernelBuilder.Build();
});

// Agent services
builder.Services.AddSingleton<AgentLogger>();
builder.Services.AddScoped<IExcelAgentService, ExcelAgentService>();
builder.Services.AddScoped<IConversationManager, ConversationManager>();
builder.Services.AddSingleton<IResponseFormatter, ResponseFormatter>();
builder.Services.AddScoped<ExportService>();

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

// Auto-detect the first responsive local LLM provider (Ollama, then LM Studio)
static async Task<(string Model, string BaseUrl)?> DetectProviderAsync()
{
    // Priority order: Ollama, then LM Studio
    var candidates = new[]
    {
        "http://localhost:11434/v1",
        "http://localhost:1234/v1",
    };

    foreach (var baseUrl in candidates)
    {
        try
        {
            using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(3) };
            var response = await httpClient.GetStringAsync($"{baseUrl}/models");
            var jsonDoc = System.Text.Json.JsonDocument.Parse(response);
            if (jsonDoc.RootElement.TryGetProperty("data", out var dataArray) && dataArray.GetArrayLength() > 0)
            {
                var firstModel = dataArray[0];
                if (firstModel.TryGetProperty("id", out var idElement))
                {
                    var model = idElement.GetString();
                    if (!string.IsNullOrEmpty(model))
                        return (model, baseUrl);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Debug("LLM provider not reachable at {BaseUrl}: {Error}", baseUrl, ex.Message);
        }
    }

    return null;
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
        Path.Combine(baseDirectory, "..", "..", "ExcelMcp.Server", fileName),
        // repo-root/src/ExcelMcp.Server/bin/Debug/net10.0/ sibling project (from ChatWeb bin output dir: up 4 = repo/src/)
        Path.Combine(baseDirectory, "..", "..", "..", "..", "ExcelMcp.Server", "bin", "Debug", "net10.0", fileName)
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
