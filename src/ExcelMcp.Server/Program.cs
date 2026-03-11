using ExcelMcp.Server.Mcp;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

// Optionally resolve a workbook path from --workbook / -w / EXCEL_MCP_WORKBOOK.
// When provided at startup, it is stored in EXCEL_MCP_WORKBOOK so all tool calls
// that omit workbook_path will fall back to it automatically.
// External MCP clients (Claude Desktop, GitHub Copilot, Cursor) pass workbook_path
// per tool call instead, so the server starts cleanly without a startup path.
var startupWorkbook = ResolveWorkbookPath(args);
if (!string.IsNullOrWhiteSpace(startupWorkbook))
{
    if (!File.Exists(startupWorkbook))
    {
        Console.Error.WriteLine($"Workbook not found at '{startupWorkbook}'.");
        return 1;
    }
    Environment.SetEnvironmentVariable("EXCEL_MCP_WORKBOOK", Path.GetFullPath(startupWorkbook));
}

var builder = Host.CreateApplicationBuilder();

// Redirect all console logging to stderr so stdout stays clean for the MCP Stdio protocol.
builder.Logging.ClearProviders();
builder.Logging.AddConsole(o => o.LogToStandardErrorThreshold = LogLevel.Trace);
builder.Logging.SetMinimumLevel(LogLevel.Warning);

builder.Services
    .AddMcpServer()
    .WithStdioServerTransport()
    .WithTools<ExcelTools>();

await builder.Build().RunAsync();
return 0;

static string? ResolveWorkbookPath(string[] arguments)
{
	if (arguments is null)
	{
		return Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
	}

	for (var i = 0; i < arguments.Length; i++)
	{
		var arg = arguments[i];
		if (string.Equals(arg, "--workbook", StringComparison.OrdinalIgnoreCase) || string.Equals(arg, "-w", StringComparison.OrdinalIgnoreCase))
		{
			if (i + 1 < arguments.Length)
			{
				return arguments[i + 1];
			}

			break;
		}

		if (arg.StartsWith("--workbook=", StringComparison.OrdinalIgnoreCase))
		{
			return arg.Substring("--workbook=".Length);
		}
	}

	return Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
}


