using ExcelMcp.ChatWeb.Logging;
using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.ChatWeb.Services;
using ExcelMcp.ChatWeb.Services.Agent;
// using ExcelMcp.ChatWeb.Services.Plugins; // Will be added when plugins are created
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

// Semantic Kernel with OpenAI chat completion
builder.Services.AddSingleton(serviceProvider =>
{
    var skOptions = serviceProvider.GetRequiredService<IOptions<SemanticKernelOptions>>().Value;
    
    var kernelBuilder = Kernel.CreateBuilder();
    
    // Add OpenAI chat completion for local LLM (LM Studio, Ollama, etc.)
    kernelBuilder.AddOpenAIChatCompletion(
        modelId: skOptions.Model,
        apiKey: skOptions.ApiKey,
        endpoint: new Uri(skOptions.BaseUrl));
    
    // Plugins will be added here when they're created
    // kernelBuilder.Plugins.AddFromObject<WorkbookStructurePlugin>();
    // kernelBuilder.Plugins.AddFromObject<WorkbookSearchPlugin>();
    // kernelBuilder.Plugins.AddFromObject<DataRetrievalPlugin>();
    
    return kernelBuilder.Build();
});

// Agent services
builder.Services.AddSingleton<AgentLogger>();
// TODO: Uncomment when implementations are created
// builder.Services.AddScoped<IExcelAgentService, ExcelAgentService>();
// builder.Services.AddScoped<IConversationManager, ConversationManager>();
// builder.Services.AddSingleton<IResponseFormatter, ResponseFormatter>();

// Plugins (will be created in Phase 3)
// builder.Services.AddSingleton<WorkbookStructurePlugin>();
// builder.Services.AddSingleton<WorkbookSearchPlugin>();
// builder.Services.AddSingleton<DataRetrievalPlugin>();

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

    var app = builder.Build();

    // Add correlation ID middleware for request tracking
    app.UseCorrelationId();

    // Add Serilog request logging
    app.UseSerilogRequestLogging();

    app.UseDefaultFiles();
    app.UseStaticFiles();

    app.MapPost("/api/chat", async (ChatRequestDto request, ChatService chatService, CancellationToken cancellationToken) =>
    {
        var response = await chatService.HandleAsync(request, cancellationToken).ConfigureAwait(false);
        return Results.Json(response);
    });

    app.MapGet("/health", () => Results.Ok(new { status = "ok" }));

    app.MapGet("/api/model", (ModelInfoDto info) => Results.Json(info));

    app.MapFallbackToFile("/index.html");

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

static void EnsureExcelMcpConfiguration(WebApplicationBuilder builder)
{
    var workbookConfigKey = $"{ExcelMcpOptions.SectionName}:WorkbookPath";
    var serverConfigKey = $"{ExcelMcpOptions.SectionName}:ServerPath";

    var workbookPath = builder.Configuration[workbookConfigKey];
    var serverPath = builder.Configuration[serverConfigKey];

    if (string.IsNullOrWhiteSpace(workbookPath))
    {
        workbookPath = Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
    }

    if (string.IsNullOrWhiteSpace(serverPath))
    {
        serverPath = Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER");
    }

    var interactive = !Console.IsInputRedirected && !Console.IsOutputRedirected && !Console.IsErrorRedirected;
    var overrides = new Dictionary<string, string?>();

    if (string.IsNullOrWhiteSpace(workbookPath) && interactive)
    {
        Console.Write("Enter the full path to the Excel workbook: ");
        workbookPath = Console.ReadLine()?.Trim();
    }

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

    if (!string.IsNullOrWhiteSpace(workbookPath))
    {
        var resolved = Path.GetFullPath(workbookPath);
        overrides[workbookConfigKey] = resolved;
        Environment.SetEnvironmentVariable("EXCEL_MCP_WORKBOOK", resolved);
    }

    if (!string.IsNullOrWhiteSpace(serverPath))
    {
        var resolved = Path.GetFullPath(serverPath);
        overrides[serverConfigKey] = resolved;
        Environment.SetEnvironmentVariable("EXCEL_MCP_SERVER", resolved);
    }

    if (overrides.Count > 0)
    {
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
