using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.ChatWeb.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;

var builder = WebApplication.CreateBuilder(args);

EnsureExcelMcpConfiguration(builder);

builder.Services.AddOptions<LlmStudioOptions>()
    .Bind(builder.Configuration.GetSection(LlmStudioOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

builder.Services.AddOptions<ExcelMcpOptions>()
    .Bind(builder.Configuration.GetSection(ExcelMcpOptions.SectionName))
    .ValidateDataAnnotations()
    .ValidateOnStart();

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

app.UseDefaultFiles();
app.UseStaticFiles();

app.MapPost("/api/chat", async (ChatRequestDto request, ChatService chatService, CancellationToken cancellationToken) =>
{
    var response = await chatService.HandleAsync(request, cancellationToken).ConfigureAwait(false);
    return Results.Json(response);
});

app.MapPost("/api/preview", async (PreviewRequestDto request, IMcpClient mcpClient, CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(request.Worksheet))
    {
        return Results.Json(new PreviewResponseDto(false, "Worksheet is required.", null, null, 0, Array.Empty<string>(), Array.Empty<PreviewRowDto>(), false, null, null));
    }

    var arguments = new JsonObject
    {
        ["worksheet"] = request.Worksheet
    };

    if (!string.IsNullOrWhiteSpace(request.Table))
    {
        arguments["table"] = request.Table;
    }

    if (request.Rows is { } rows && rows > 0)
    {
        arguments["rows"] = rows;
    }

    if (!string.IsNullOrWhiteSpace(request.Cursor))
    {
        arguments["cursor"] = request.Cursor;
    }

    try
    {
        var result = await mcpClient.CallToolAsync("excel-preview-table", arguments, cancellationToken).ConfigureAwait(false);
        var textContent = result.Content.FirstOrDefault(static c => string.Equals(c.Type, "text", StringComparison.OrdinalIgnoreCase))?.Text;
        var jsonContent = result.Content.FirstOrDefault(static c => string.Equals(c.Type, "json", StringComparison.OrdinalIgnoreCase))?.Json as JsonObject;

        if (result.IsError)
        {
            var errorText = string.IsNullOrWhiteSpace(textContent) ? "Preview tool reported an error." : textContent;
            return Results.Json(new PreviewResponseDto(false, errorText, null, null, 0, Array.Empty<string>(), Array.Empty<PreviewRowDto>(), false, null, null));
        }

        if (jsonContent is null)
        {
            var message = string.IsNullOrWhiteSpace(textContent) ? "Preview tool did not return structured data." : textContent;
            return Results.Json(new PreviewResponseDto(false, message, null, null, 0, Array.Empty<string>(), Array.Empty<PreviewRowDto>(), false, null, null));
        }

        var payload = jsonContent.Deserialize<PreviewToolPayload>(new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (payload is null)
        {
            return Results.Json(new PreviewResponseDto(false, "Unable to parse preview payload.", null, null, 0, Array.Empty<string>(), Array.Empty<PreviewRowDto>(), false, null, null));
        }

        var headers = payload.Headers ?? Array.Empty<string>();
        var rowsPayload = payload.Rows ?? Array.Empty<PreviewToolRowPayload>();
        var rowsResult = rowsPayload
            .Select(row => new PreviewRowDto(row.RowNumber, row.Values ?? Array.Empty<string?>()))
            .ToArray();

        return Results.Json(new PreviewResponseDto(
            true,
            null,
            payload.Worksheet,
            payload.Table,
            payload.Offset,
            headers,
            rowsResult,
            payload.HasMore,
            payload.NextCursor,
            textContent
        ));
    }
    catch (Exception ex)
    {
        return Results.Json(new PreviewResponseDto(false, ex.Message, null, null, 0, Array.Empty<string>(), Array.Empty<PreviewRowDto>(), false, null, null));
    }
});

app.MapGet("/api/resources", async (IMcpClient mcpClient, CancellationToken cancellationToken) =>
{
    try
    {
        var resources = await mcpClient.ListResourcesAsync(cancellationToken).ConfigureAwait(false);
        var worksheetNames = resources
            .Where(static r => string.Equals(r.Uri.Scheme, "excel", StringComparison.OrdinalIgnoreCase)
                && string.Equals(r.Uri.Host, "worksheet", StringComparison.OrdinalIgnoreCase))
            .Select(static r => Uri.UnescapeDataString(r.Uri.AbsolutePath.Trim('/').Split('/', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty))
            .Where(static name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        return Results.Json(new { worksheets = worksheetNames });
    }
    catch (Exception ex)
    {
        return Results.Json(new { worksheets = Array.Empty<string>(), error = ex.Message });
    }
});

app.MapGet("/health", () => Results.Ok(new { status = "ok" }));

app.MapGet("/api/model", (ModelInfoDto info) => Results.Json(info));

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
