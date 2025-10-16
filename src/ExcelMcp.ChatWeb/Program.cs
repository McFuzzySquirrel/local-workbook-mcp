using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.ChatWeb.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;

var builder = WebApplication.CreateBuilder(args);

EnsureExcelMcpConfiguration(builder);

builder.Services.AddOptions<OllamaOptions>()
    .Bind(builder.Configuration.GetSection(OllamaOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddOptions<ExcelMcpOptions>()
    .Bind(builder.Configuration.GetSection(ExcelMcpOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddSingleton<McpClientHost>();
builder.Services.AddSingleton<IMcpClient>(static sp => sp.GetRequiredService<McpClientHost>());
builder.Services.AddSingleton<IHostedService>(static sp => sp.GetRequiredService<McpClientHost>());

builder.Services.AddHttpClient<IOllamaClient, OllamaClient>((serviceProvider, client) =>
{
    var options = serviceProvider.GetRequiredService<IOptions<OllamaOptions>>().Value;
    client.BaseAddress = new Uri(options.BaseUrl, UriKind.Absolute);
    client.Timeout = TimeSpan.FromSeconds(120);
});

builder.Services.AddSingleton<ChatService>();

var app = builder.Build();

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapPost("/api/chat", async (ChatRequestDto request, ChatService chatService, CancellationToken cancellationToken) =>
{
    var response = await chatService.HandleAsync(request, cancellationToken).ConfigureAwait(false);
    return Results.Json(response);
});

app.MapGet("/health", () => Results.Ok(new { status = "ok" }));

app.MapFallbackToFile("/index.html");

app.Run();

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
