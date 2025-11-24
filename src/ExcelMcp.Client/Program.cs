using System.Text.Json;
using System.Text.Json.Nodes;
using ExcelMcp.Client.Mcp;

if (args.Length == 0 || args.Contains("--help") || args.Contains("-h"))
{
    ShowHelp();
    return 0;
}

var workbookPath = ResolveArgument(args, "--workbook", "-w") ?? Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
var serverPath = ResolveArgument(args, "--server", "-s") ?? Environment.GetEnvironmentVariable("EXCEL_MCP_SERVER");

// Interactive prompt for workbook if missing
if (string.IsNullOrWhiteSpace(workbookPath))
{
    Console.Write("Enter the full path to the Excel workbook: ");
    workbookPath = Console.ReadLine()?.Trim();
}

if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
{
    Console.Error.WriteLine($"Error: Workbook not found at '{workbookPath}'");
    return 1;
}

// Interactive prompt for server if missing
if (string.IsNullOrWhiteSpace(serverPath))
{
    // Try to find it relative to current execution
    serverPath = FindServerExecutable();
    
    if (string.IsNullOrWhiteSpace(serverPath))
    {
        Console.Write("Enter the full path to the Excel MCP server executable: ");
        serverPath = Console.ReadLine()?.Trim();
    }
}

if (string.IsNullOrWhiteSpace(serverPath) || !File.Exists(serverPath))
{
    Console.Error.WriteLine($"Error: Server executable not found at '{serverPath}'");
    return 1;
}

// Determine command
var command = args.FirstOrDefault(a => !a.StartsWith("-"));
if (string.IsNullOrEmpty(command))
{
    command = "list"; // Default
}

try
{
    await using var client = new McpProcessClient(serverPath, workbookPath);
    await client.InitializeAsync(CancellationToken.None);

    switch (command.ToLowerInvariant())
    {
        case "list":
            await ListToolsAsync(client);
            break;
        case "resources":
            await ListResourcesAsync(client);
            break;
        case "search":
            var query = args.SkipWhile(a => a != "search").Skip(1).FirstOrDefault();
            if (string.IsNullOrEmpty(query))
            {
                Console.Error.WriteLine("Error: Search query required.");
                return 1;
            }
            await SearchAsync(client, query);
            break;
        case "preview":
            var sheet = args.SkipWhile(a => a != "preview").Skip(1).FirstOrDefault();
            if (string.IsNullOrEmpty(sheet))
            {
                Console.Error.WriteLine("Error: Worksheet name required.");
                return 1;
            }
            var rowsStr = ResolveArgument(args, "--rows");
            int rows = int.TryParse(rowsStr, out var r) ? r : 10;
            await PreviewAsync(client, sheet, rows);
            break;
        default:
            Console.Error.WriteLine($"Unknown command: {command}");
            ShowHelp();
            return 1;
    }
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}

return 0;

static void ShowHelp()
{
    Console.WriteLine("Excel MCP Client");
    Console.WriteLine("Usage: ExcelMcp.Client [options] <command> [args]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  --workbook, -w <path>   Path to Excel workbook (or set EXCEL_MCP_WORKBOOK)");
    Console.WriteLine("  --server, -s <path>     Path to MCP server executable (or set EXCEL_MCP_SERVER)");
    Console.WriteLine();
    Console.WriteLine("Commands:");
    Console.WriteLine("  list                    List available tools");
    Console.WriteLine("  resources               List available resources");
    Console.WriteLine("  search <query>          Search the workbook");
    Console.WriteLine("  preview <sheet>         Preview rows from a sheet");
    Console.WriteLine("    --rows <n>            Number of rows to preview (default: 10)");
}

static string? ResolveArgument(string[] args, string longName, string? shortName = null)
{
    for (int i = 0; i < args.Length; i++)
    {
        if (args[i] == longName || (shortName != null && args[i] == shortName))
        {
            if (i + 1 < args.Length) return args[i + 1];
        }
        if (args[i].StartsWith($"{longName}="))
        {
            return args[i].Substring(longName.Length + 1);
        }
    }
    return null;
}

static string? FindServerExecutable()
{
    var baseDir = AppContext.BaseDirectory;
    var candidates = new[]
    {
        Path.Combine(baseDir, "ExcelMcp.Server.exe"),
        Path.Combine(baseDir, "ExcelMcp.Server"),
        Path.Combine(baseDir, "..", "ExcelMcp.Server", "ExcelMcp.Server.exe"), // Dev layout
        Path.Combine(baseDir, "..", "ExcelMcp.Server", "bin", "Debug", "net9.0", "ExcelMcp.Server.exe") // Dev layout
    };

    foreach (var path in candidates)
    {
        if (File.Exists(path)) return Path.GetFullPath(path);
    }
    return null;
}

static async Task ListToolsAsync(McpProcessClient client)
{
    var tools = await client.ListToolsAsync(CancellationToken.None);
    Console.WriteLine($"Found {tools.Count} tools:");
    foreach (var tool in tools)
    {
        Console.WriteLine($"- {tool.Name}: {tool.Description}");
    }
}

static async Task ListResourcesAsync(McpProcessClient client)
{
    var resources = await client.ListResourcesAsync(CancellationToken.None);
    Console.WriteLine($"Found {resources.Count} resources:");
    foreach (var res in resources)
    {
        Console.WriteLine($"- {res.Name} ({res.MimeType})");
        Console.WriteLine($"  URI: {res.Uri}");
    }
}

static async Task SearchAsync(McpProcessClient client, string query)
{
    var args = new JsonObject
    {
        ["query"] = query
    };
    
    var result = await client.CallToolAsync("excel-search", args, CancellationToken.None);
    if (result.IsError)
    {
        Console.WriteLine("Search failed.");
    }
    
    foreach (var content in result.Content)
    {
        Console.WriteLine(content.Text);
    }
}

static async Task PreviewAsync(McpProcessClient client, string sheet, int rows)
{
    var args = new JsonObject
    {
        ["worksheet"] = sheet,
        ["rows"] = rows
    };
    
    var result = await client.CallToolAsync("excel-preview-table", args, CancellationToken.None);
    if (result.IsError)
    {
        Console.WriteLine("Preview failed.");
    }
    
    foreach (var content in result.Content)
    {
        Console.WriteLine(content.Text);
    }
}
