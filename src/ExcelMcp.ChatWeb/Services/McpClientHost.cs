using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.Client.Mcp;
using Microsoft.Extensions.Options;

namespace ExcelMcp.ChatWeb.Services;

public sealed class McpClientHost : IHostedService, IAsyncDisposable, IMcpClient
{
    private readonly ExcelMcpOptions _options;
    private McpProcessClient? _client;

    public McpClientHost(IOptions<ExcelMcpOptions> options)
    {
        _options = options.Value;
    }

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        var workbookPathSetting = string.IsNullOrWhiteSpace(_options.WorkbookPath)
            ? Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK")
            : _options.WorkbookPath;

        if (string.IsNullOrWhiteSpace(workbookPathSetting))
        {
            throw new InvalidOperationException("ExcelMcp:WorkbookPath must be configured or set via EXCEL_MCP_WORKBOOK.");
        }

        var serverPathSetting = string.IsNullOrWhiteSpace(_options.ServerPath)
            ? Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER")
            : _options.ServerPath;

        if (string.IsNullOrWhiteSpace(serverPathSetting))
        {
            throw new InvalidOperationException("ExcelMcp:ServerPath must be configured or set via EXCEL_MCP_SERVER.");
        }

        var workbookPath = Path.GetFullPath(workbookPathSetting);
        var serverPath = Path.GetFullPath(serverPathSetting);

        if (!File.Exists(serverPath))
        {
            throw new FileNotFoundException("Unable to locate MCP server executable.", serverPath);
        }

        _client = new McpProcessClient(serverPath, workbookPath);
        await _client.InitializeAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task StopAsync(CancellationToken cancellationToken)
    {
        if (_client is not null)
        {
            await _client.DisposeAsync().ConfigureAwait(false);
            _client = null;
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (_client is not null)
        {
            await _client.DisposeAsync().ConfigureAwait(false);
            _client = null;
        }
    }

    public async Task<IReadOnlyList<McpToolDefinition>> ListToolsAsync(CancellationToken cancellationToken)
    {
        return await GetClient().ListToolsAsync(cancellationToken).ConfigureAwait(false);
    }

    public async Task<McpToolCallResult> CallToolAsync(string name, JsonNode? arguments, CancellationToken cancellationToken)
    {
        return await GetClient().CallToolAsync(name, arguments, cancellationToken).ConfigureAwait(false);
    }

    public async Task<IReadOnlyList<McpResource>> ListResourcesAsync(CancellationToken cancellationToken)
    {
        return await GetClient().ListResourcesAsync(cancellationToken).ConfigureAwait(false);
    }

    private McpProcessClient GetClient()
    {
        if (_client is null)
        {
            throw new InvalidOperationException("MCP client is not initialized.");
        }

        return _client;
    }
}
