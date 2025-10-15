using System.Diagnostics;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelMcp.Client.Mcp;

internal sealed class McpProcessClient : IAsyncDisposable
{
    private readonly Process _process;
    private readonly JsonRpcClient _jsonRpc;
    private bool _initialized;

    public McpProcessClient(string serverPath, string workbookPath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = serverPath,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };

        startInfo.ArgumentList.Add("--workbook");
        startInfo.ArgumentList.Add(workbookPath);

        _process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start MCP server process.");
        _process.EnableRaisingEvents = true;

        _jsonRpc = new JsonRpcClient(_process.StandardOutput.BaseStream, _process.StandardInput.BaseStream);

        _ = Task.Run(async () =>
        {
            var reader = _process.StandardError;
            while (!reader.EndOfStream)
            {
                var line = await reader.ReadLineAsync().ConfigureAwait(false);
                if (!string.IsNullOrWhiteSpace(line))
                {
                    Console.Error.WriteLine($"[server] {line}");
                }
            }
        });
    }

    public async Task InitializeAsync(CancellationToken cancellationToken)
    {
        var parameters = new JsonObject
        {
            ["clientInfo"] = new JsonObject
            {
                ["name"] = "excel-mcp-client",
                ["version"] = "0.1.0"
            },
            ["capabilities"] = new JsonObject()
        };

        using var response = await _jsonRpc.SendRequestAsync("initialize", parameters, cancellationToken).ConfigureAwait(false);
        _initialized = true;
    }

    public async Task<IReadOnlyList<McpToolDefinition>> ListToolsAsync(CancellationToken cancellationToken)
    {
        using var response = await _jsonRpc.SendRequestAsync("tools/list", new JsonObject(), cancellationToken).ConfigureAwait(false);
        var result = response.RootElement.GetProperty("result").GetProperty("tools");
        var tools = new List<McpToolDefinition>();
        foreach (var element in result.EnumerateArray())
        {
            var name = element.GetProperty("name").GetString()!;
            var description = element.TryGetProperty("description", out var descriptionElement) ? descriptionElement.GetString() ?? string.Empty : string.Empty;
            var schema = JsonNode.Parse(element.GetProperty("inputSchema").GetRawText())!;
            tools.Add(new McpToolDefinition(name, description, schema));
        }

        return tools;
    }

    public async Task<McpToolCallResult> CallToolAsync(string name, JsonNode? arguments, CancellationToken cancellationToken)
    {
        var parameters = new JsonObject
        {
            ["name"] = name,
            ["arguments"] = arguments
        };

        using var response = await _jsonRpc.SendRequestAsync("tools/call", parameters, cancellationToken).ConfigureAwait(false);
        var result = response.RootElement.GetProperty("result");
        var contentItems = new List<McpContentItem>();

        foreach (var item in result.GetProperty("content").EnumerateArray())
        {
            var type = item.GetProperty("type").GetString();
            switch (type)
            {
                case "text":
                    contentItems.Add(new McpTextContent(item.GetProperty("text").GetString() ?? string.Empty));
                    break;
                case "json":
                    var json = JsonNode.Parse(item.GetProperty("json").GetRawText()) ?? new JsonObject();
                    contentItems.Add(new McpJsonContent(json));
                    break;
            }
        }

        var isError = result.TryGetProperty("isError", out var isErrorElement) && isErrorElement.GetBoolean();
        return new McpToolCallResult(contentItems, isError);
    }

    public async Task<IReadOnlyList<McpResource>> ListResourcesAsync(CancellationToken cancellationToken)
    {
        using var response = await _jsonRpc.SendRequestAsync("resources/list", new JsonObject(), cancellationToken).ConfigureAwait(false);
        var result = response.RootElement.GetProperty("result").GetProperty("resources");
        var resources = new List<McpResource>();
        foreach (var element in result.EnumerateArray())
        {
            var uri = new Uri(element.GetProperty("uri").GetString()!);
            var name = element.GetProperty("name").GetString() ?? uri.ToString();
            var description = element.TryGetProperty("description", out var descriptionElement) ? descriptionElement.GetString() : null;
            var mimeType = element.TryGetProperty("mimeType", out var mimeElement) ? mimeElement.GetString() : null;
            resources.Add(new McpResource(uri, name, description, mimeType));
        }

        return resources;
    }

    public async Task<McpResourceContent> ReadResourceAsync(Uri uri, CancellationToken cancellationToken)
    {
        var parameters = new JsonObject
        {
            ["uri"] = uri.ToString()
        };

        using var response = await _jsonRpc.SendRequestAsync("resources/read", parameters, cancellationToken).ConfigureAwait(false);
        var contents = response.RootElement.GetProperty("result").GetProperty("contents");
        var first = contents.EnumerateArray().First();
        var text = first.TryGetProperty("text", out var textElement) ? textElement.GetString() : null;
        var mimeType = first.TryGetProperty("mimeType", out var mimeElement) ? mimeElement.GetString() : null;
        return new McpResourceContent(uri, mimeType, text);
    }

    public async ValueTask DisposeAsync()
    {
        if (_initialized)
        {
            try
            {
                await _jsonRpc.SendRequestAsync("shutdown", new JsonObject(), CancellationToken.None).ConfigureAwait(false);
            }
            catch
            {
                // ignore
            }
        }

        await _jsonRpc.DisposeAsync().ConfigureAwait(false);

        if (!_process.HasExited)
        {
            if (!_process.WaitForExit(2000))
            {
                _process.Kill(entireProcessTree: true);
                _process.WaitForExit();
            }
        }

        _process.Dispose();
    }
}
