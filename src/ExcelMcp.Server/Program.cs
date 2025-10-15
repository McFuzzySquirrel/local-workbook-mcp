using ExcelMcp.Server.Excel;
using ExcelMcp.Server.Mcp;

var workbookPath = ResolveWorkbookPath(args);
if (workbookPath is null)
{
	Console.Error.WriteLine("A workbook path must be provided via --workbook or the EXCEL_MCP_WORKBOOK environment variable.");
	return 1;
}

if (!File.Exists(workbookPath))
{
	Console.Error.WriteLine($"Workbook not found at '{workbookPath}'.");
	return 1;
}

using var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, eventArgs) =>
{
	eventArgs.Cancel = true;
	cts.Cancel();
};

var service = new ExcelWorkbookService(workbookPath);
var server = new McpServer(service);

await server.RunAsync(cts.Token);
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
