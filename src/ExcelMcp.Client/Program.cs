using System.Linq;
using System.Text.Json.Nodes;
using ExcelMcp.Client.Mcp;

var cancellationTokenSource = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) =>
{
	eventArgs.Cancel = true;
	cancellationTokenSource.Cancel();
};

var (globalOptions, command, commandArgs) = ParseArguments(args);
if (command is null)
{
	PrintUsage();
	return 1;
}

var workbookPath = ResolveWorkbookPath(globalOptions);
if (workbookPath is null)
{
	Console.Error.WriteLine("A workbook path is required. Use --workbook or set EXCEL_MCP_WORKBOOK.");
	return 1;
}

var serverPath = ResolveServerPath(globalOptions);
if (serverPath is null || !File.Exists(serverPath))
{
	Console.Error.WriteLine("Unable to locate the MCP server executable. Specify --server or set EXCEL_MCP_SERVER.");
	return 1;
}

await using var client = new McpProcessClient(serverPath, workbookPath);
await client.InitializeAsync(cancellationTokenSource.Token);

switch (command)
{
	case "list":
		await HandleListAsync(client, cancellationTokenSource.Token);
		break;
	case "search":
		await HandleSearchAsync(client, commandArgs, cancellationTokenSource.Token);
		break;
	case "preview":
		await HandlePreviewAsync(client, commandArgs, cancellationTokenSource.Token);
		break;
	case "resources":
		await HandleResourcesAsync(client, cancellationTokenSource.Token);
		break;
	default:
		Console.Error.WriteLine($"Unknown command '{command}'.");
		PrintUsage();
		return 1;
}

return 0;

static async Task HandleListAsync(McpProcessClient client, CancellationToken cancellationToken)
{
	var result = await client.CallToolAsync("excel.list_structure", new JsonObject(), cancellationToken);
	foreach (var content in result.Content)
	{
		switch (content)
		{
			case McpTextContent text:
				Console.WriteLine(text.Text);
				break;
			case McpJsonContent json:
				Console.WriteLine(json.Json.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
				break;
		}
	}
}

static async Task HandleSearchAsync(McpProcessClient client, Queue<string> argsQueue, CancellationToken cancellationToken)
{
	var options = ParseOptionPairs(argsQueue);
	if (!options.TryGetValue("--query", out var query) && !options.TryGetValue("-q", out query))
	{
		Console.Error.WriteLine("The --query option is required for search.");
		return;
	}

	var payload = new JsonObject
	{
		["query"] = query
	};

	if (options.TryGetValue("--worksheet", out var worksheet) || options.TryGetValue("-w", out worksheet))
	{
		payload["worksheet"] = worksheet;
	}

	if (options.TryGetValue("--table", out var table))
	{
		payload["table"] = table;
	}

	if (options.TryGetValue("--limit", out var limitText) && int.TryParse(limitText, out var limit) && limit > 0)
	{
		payload["limit"] = limit;
	}

	if (options.ContainsKey("--case-sensitive"))
	{
		payload["caseSensitive"] = true;
	}

	var result = await client.CallToolAsync("excel.search", payload, cancellationToken);
	foreach (var content in result.Content)
	{
		switch (content)
		{
			case McpJsonContent json:
				Console.WriteLine(json.Json.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
				break;
			case McpTextContent text:
				Console.WriteLine(text.Text);
				break;
		}
	}
}

static async Task HandlePreviewAsync(McpProcessClient client, Queue<string> argsQueue, CancellationToken cancellationToken)
{
	var options = ParseOptionPairs(argsQueue);
	if (!options.TryGetValue("--worksheet", out var worksheet) && !options.TryGetValue("-w", out worksheet))
	{
		Console.Error.WriteLine("The --worksheet option is required for preview.");
		return;
	}

	var payload = new JsonObject
	{
		["worksheet"] = worksheet
	};

	if (options.TryGetValue("--table", out var table) || options.TryGetValue("-t", out table))
	{
		payload["table"] = table;
	}

	if (options.TryGetValue("--rows", out var rowsText) && int.TryParse(rowsText, out var rows) && rows > 0)
	{
		payload["rows"] = rows;
	}

	var result = await client.CallToolAsync("excel.preview_table", payload, cancellationToken);
	foreach (var content in result.Content.OfType<McpTextContent>())
	{
		Console.WriteLine(content.Text);
	}
}

static async Task HandleResourcesAsync(McpProcessClient client, CancellationToken cancellationToken)
{
	var resources = await client.ListResourcesAsync(cancellationToken);
	foreach (var resource in resources)
	{
		Console.WriteLine($"{resource.Uri} | {resource.Name} | {resource.Description}");
	}
}

static (Dictionary<string, string?> GlobalOptions, string? Command, Queue<string> CommandArgs) ParseArguments(string[] arguments)
{
	var options = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
	var queue = new Queue<string>(arguments ?? Array.Empty<string>());

	while (queue.Count > 0)
	{
		var token = queue.Peek();
		if (!token.StartsWith("-"))
		{
			break;
		}

		queue.Dequeue();
		if (token.Contains('='))
		{
			var split = token.Split('=', 2);
			options[split[0]] = split[1];
			continue;
		}

		if (queue.Count > 0 && !queue.Peek().StartsWith("-"))
		{
			options[token] = queue.Dequeue();
			continue;
		}

		options[token] = "true";
	}

	var command = queue.Count > 0 ? queue.Dequeue() : null;
	return (options, command, new Queue<string>(queue));
}

static Dictionary<string, string?> ParseOptionPairs(Queue<string> queue)
{
	var options = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
	while (queue.Count > 0)
	{
		var token = queue.Peek();
		if (!token.StartsWith("-"))
		{
			break;
		}

		queue.Dequeue();

		if (token.Contains('='))
		{
			var split = token.Split('=', 2);
			options[split[0]] = split[1];
			continue;
		}

		if (queue.Count > 0 && !queue.Peek().StartsWith("-"))
		{
			options[token] = queue.Dequeue();
			continue;
		}

		options[token] = "true";
	}

	return options;
}

static string? ResolveWorkbookPath(Dictionary<string, string?> options)
{
	if (options.TryGetValue("--workbook", out var workbook) || options.TryGetValue("-w", out workbook))
	{
		return string.IsNullOrWhiteSpace(workbook) ? null : Path.GetFullPath(workbook);
	}

	var env = Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
	return string.IsNullOrWhiteSpace(env) ? null : Path.GetFullPath(env);
}

static string? ResolveServerPath(Dictionary<string, string?> options)
{
	if (options.TryGetValue("--server", out var server))
	{
		return string.IsNullOrWhiteSpace(server) ? null : Path.GetFullPath(server);
	}

	var env = Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER");
	if (!string.IsNullOrWhiteSpace(env))
	{
		return Path.GetFullPath(env);
	}

	var defaultPath = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "ExcelMcp.Server", "bin", "Debug", "net9.0", OperatingSystem.IsWindows() ? "ExcelMcp.Server.exe" : "ExcelMcp.Server");
	defaultPath = Path.GetFullPath(defaultPath);
	return File.Exists(defaultPath) ? defaultPath : null;
}

static void PrintUsage()
{
	Console.WriteLine("Usage: ExcelMcp.Client --workbook <path> [--server <path>] <command> [options]");
	Console.WriteLine();
	Console.WriteLine("Commands:");
	Console.WriteLine("  list                         Summarize workbook structure using the MCP server.");
	Console.WriteLine("  search --query <text>        Search workbook content (options: --worksheet, --table, --limit, --case-sensitive).");
	Console.WriteLine("  preview --worksheet <name>   Preview worksheet or table as CSV (options: --table, --rows).");
	Console.WriteLine("  resources                    List exposed MCP resources.");
}
