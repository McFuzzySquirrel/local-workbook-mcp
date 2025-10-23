using System.Text.Json.Nodes;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.Client.Mcp;
using Microsoft.Extensions.Options;

namespace ExcelMcp.ChatWeb.Services;

public sealed class McpClientHost : IHostedService, IAsyncDisposable, IMcpClient
{
    private readonly ExcelMcpOptions _options;
    private McpProcessClient? _client;
    private string? _currentWorkbookPath;
    private readonly SemaphoreSlim _initLock = new(1, 1);

    public McpClientHost(IOptions<ExcelMcpOptions> options)
    {
        _options = options.Value;
    }

    public bool IsInitialized => _client != null;

    public string? CurrentWorkbookPath => _currentWorkbookPath;

    public async Task StartAsync(CancellationToken cancellationToken)
    {
        // Don't auto-initialize - wait for user to select workbook via InitializeWithWorkbookAsync
        // Only initialize if paths are explicitly configured
        var workbookPathSetting = string.IsNullOrWhiteSpace(_options.WorkbookPath)
            ? Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK")
            : _options.WorkbookPath;

        if (!string.IsNullOrWhiteSpace(workbookPathSetting))
        {
            var serverPathSetting = string.IsNullOrWhiteSpace(_options.ServerPath)
                ? Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER")
                : _options.ServerPath;

            if (!string.IsNullOrWhiteSpace(serverPathSetting))
            {
                await InitializeWithWorkbookAsync(workbookPathSetting, serverPathSetting, cancellationToken);
            }
        }
    }

    /// <summary>
    /// Initializes or reinitializes the MCP client with a new workbook.
    /// </summary>
    public async Task InitializeWithWorkbookAsync(string workbookPath, string? serverPath = null, CancellationToken cancellationToken = default)
    {
        await _initLock.WaitAsync(cancellationToken);
        try
        {
            // Dispose existing client if reinitializing
            if (_client != null)
            {
                await _client.DisposeAsync();
                _client = null;
            }

            var resolvedServerPath = serverPath ?? _options.ServerPath;
            
            if (string.IsNullOrWhiteSpace(resolvedServerPath))
            {
                resolvedServerPath = Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER") 
                    ?? "src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server.exe";
            }

            var fullWorkbookPath = Path.GetFullPath(workbookPath);
            var fullServerPath = Path.GetFullPath(resolvedServerPath);

            if (!File.Exists(fullWorkbookPath))
            {
                throw new FileNotFoundException($"Workbook file not found: {fullWorkbookPath}");
            }

            if (!File.Exists(fullServerPath))
            {
                throw new FileNotFoundException("Unable to locate MCP server executable.", fullServerPath);
            }

            _client = new McpProcessClient(fullServerPath, fullWorkbookPath);
            await _client.InitializeAsync(cancellationToken).ConfigureAwait(false);
            _currentWorkbookPath = fullWorkbookPath;
        }
        finally
        {
            _initLock.Release();
        }
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

    private McpProcessClient GetClient()
    {
        if (_client is null)
        {
            throw new InvalidOperationException("MCP client is not initialized.");
        }

        return _client;
    }
}
