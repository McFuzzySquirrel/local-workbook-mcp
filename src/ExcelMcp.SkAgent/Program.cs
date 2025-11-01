using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Spectre.Console;
using ExcelMcp.SkAgent;

AnsiConsole.Clear();

var workbookPath = ResolveWorkbookPath(args);
if (string.IsNullOrWhiteSpace(workbookPath))
{
    workbookPath = AnsiConsole.Ask<string>("[green]Enter path to Excel workbook:[/]");
}

if (!File.Exists(workbookPath))
{
    AnsiConsole.MarkupLine($"[red]Error: Workbook not found at '{workbookPath}'[/]");
    return 1;
}

var config = LoadConfiguration();
if (config is null)
{
    config = new AgentConfiguration
    {
        BaseUrl = "http://localhost:1234/v1",
        ModelId = "local-model",
        ApiKey = "lm-studio"
    };
}

// Try to detect actual running model
string actualModel = await DetectRunningModel(config.BaseUrl);

using var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) =>
{
    e.Cancel = true;
    cts.Cancel();
};

var agent = new ExcelAgent(workbookPath, config);
await agent.InitializeAsync(cts.Token);

var history = new ChatHistory();
history.AddSystemMessage(@"You are an Excel workbook assistant. You have access to function tools that let you read data from the workbook.

CRITICAL INSTRUCTIONS - YOU MUST FOLLOW THESE:

1. NEVER say ""I don't have access to the data"" - YOU DO! Use the tools!
2. When asked about data, rows, or content - IMMEDIATELY call the appropriate tool
3. ALWAYS use tools to get real data - never make assumptions or refuse to look
4. The data is RIGHT THERE in the workbook - just use the tools to read it!

YOUR AVAILABLE TOOLS:

1. get_workbook_summary - Gets metadata about the workbook (worksheet names, row counts, columns)
   Use this FIRST if you don't know what's in the workbook!

2. list_structure - Lists all worksheets and tables with their column names
   Use this to understand the structure before querying data

3. preview_table - Gets actual row data from a worksheet or table
   Parameters: worksheet (required), table (optional), rows (default 20)
   Use this to see actual data!

4. search - Searches for specific text across all sheets
   Use this to find data containing specific values

WORKFLOW - FOLLOW THIS PATTERN:

Step 1: If you don't know the structure → Call get_workbook_summary or list_structure
Step 2: Once you know worksheet names → Call preview_table to get actual data
Step 3: Analyze the real data you retrieved and answer the question

EXAMPLES OF CORRECT BEHAVIOR:

User: ""what data is in this workbook?""
✅ Call get_workbook_summary → See worksheet names → Describe what you found

User: ""what's in the first sheet?""
✅ Call list_structure → Get first worksheet name → Call preview_table on it

User: ""show me the data""
✅ Call list_structure → Get worksheet name → Call preview_table(worksheet='SheetName', rows=20)

User: ""are there any high priority items?""
✅ Call preview_table to get data → Look through it → Report high priority items OR search for ""high"" OR ""priority""

WRONG BEHAVIOR - NEVER DO THIS:

❌ ""I don't have access to the data""
❌ ""I cannot retrieve the data""
❌ ""The tools don't allow me to see that""
❌ Making up data without calling tools

REMEMBER: You CAN see the data. Just call the tools! Start with get_workbook_summary if unsure.");


// Render initial display
AnsiConsole.Clear();
RenderHeader(workbookPath, config.ModelId, actualModel);

while (!cts.Token.IsCancellationRequested)
{
    // Show prompt
    AnsiConsole.Markup($"\n[green]>[/] ");
    var prompt = Console.ReadLine();
    
    if (string.IsNullOrWhiteSpace(prompt))
    {
        continue;
    }

    var command = prompt.Trim().ToLowerInvariant();
    
    if (command is "exit" or "quit" or "q")
    {
        break;
    }

    if (command is "clear" or "cls")
    {
        AnsiConsole.Clear();
        RenderHeader(workbookPath, config.ModelId, actualModel);
        continue;
    }

    if (command is "help" or "?")
    {
        ShowHelp();
        continue;
    }

    // Workbook switching command
    if (command.StartsWith("load ") || command.StartsWith("open ") || command is "load" or "open" or "switch")
    {
        string? newPath = null;
        
        if (command.Contains(' '))
        {
            newPath = prompt.Substring(prompt.IndexOf(' ') + 1).Trim().Trim('"');
        }
        else
        {
            newPath = AnsiConsole.Ask<string>("[green]Enter path to new workbook:[/]");
        }

        if (!File.Exists(newPath))
        {
            AnsiConsole.MarkupLine($"[red]Error: Workbook not found at '{newPath}'[/]");
            AnsiConsole.WriteLine();
            continue;
        }

        workbookPath = newPath;
        
        // Reinitialize the agent with the new workbook
        try
        {
            agent = new ExcelAgent(workbookPath, config);
            await agent.InitializeAsync(cts.Token);
            
            // Clear chat history for new workbook
            history = new ChatHistory();
            history.AddSystemMessage(@"You are an Excel workbook assistant. You have access to function tools that let you read data from the workbook.

CRITICAL INSTRUCTIONS - YOU MUST FOLLOW THESE:

1. NEVER say ""I don't have access to the data"" - YOU DO! Use the tools!
2. When asked about data, rows, or content - IMMEDIATELY call the appropriate tool
3. ALWAYS use tools to get real data - never make assumptions or refuse to look
4. The data is RIGHT THERE in the workbook - just use the tools to read it!

YOUR AVAILABLE TOOLS:

1. get_workbook_summary - Gets metadata about the workbook (worksheet names, row counts, columns)
   Use this FIRST if you don't know what's in the workbook!

2. list_structure - Lists all worksheets and tables with their column names
   Use this to understand the structure before querying data

3. preview_table - Gets actual row data from a worksheet or table
   Parameters: worksheet (required), table (optional), rows (default 20)
   Use this to see actual data!

4. search - Searches for specific text across all sheets
   Use this to find data containing specific values

WORKFLOW - FOLLOW THIS PATTERN:

Step 1: If you don't know the structure → Call get_workbook_summary or list_structure
Step 2: Once you know worksheet names → Call preview_table to get actual data
Step 3: Analyze the real data you retrieved and answer the question

EXAMPLES OF CORRECT BEHAVIOR:

User: ""what data is in this workbook?""
✅ Call get_workbook_summary → See worksheet names → Describe what you found

User: ""what's in the first sheet?""
✅ Call list_structure → Get first worksheet name → Call preview_table on it

User: ""show me the data""
✅ Call list_structure → Get worksheet name → Call preview_table(worksheet='SheetName', rows=20)

User: ""are there any high priority items?""
✅ Call preview_table to get data → Look through it → Report high priority items OR search for ""high"" OR ""priority""

WRONG BEHAVIOR - NEVER DO THIS:

❌ ""I don't have access to the data""
❌ ""I cannot retrieve the data""
❌ ""The tools don't allow me to see that""
❌ Making up data without calling tools

REMEMBER: You CAN see the data. Just call the tools! Start with get_workbook_summary if unsure.");

            AnsiConsole.Clear();
            RenderHeader(workbookPath, config.ModelId, actualModel);
            
            var successPanel = new Panel($"[green bold]✓[/] Successfully loaded [yellow bold]{Path.GetFileName(workbookPath)}[/]")
                .Border(BoxBorder.Rounded)
                .BorderColor(Color.Green)
                .Padding(0, 0);
            AnsiConsole.Write(successPanel);
            AnsiConsole.WriteLine();
        }
        catch (Exception ex)
        {
            var errorPanel = new Panel($"[red bold]Error:[/] [red]{ex.Message.EscapeMarkup()}[/]\n[yellow]Tip: Make sure your LLM server is running and accessible.[/]")
                .Border(BoxBorder.Rounded)
                .BorderColor(Color.Red)
                .Padding(0, 0);
            AnsiConsole.Write(errorPanel);
            AnsiConsole.WriteLine();
        }
        
        continue;
    }

    history.AddUserMessage(prompt);

    try
    {
        await AnsiConsole.Status()
            .Spinner(Spinner.Known.Dots)
            .SpinnerStyle(Style.Parse("green"))
            .StartAsync("[yellow]Thinking...[/]", async ctx =>
            {
                await agent.ProcessMessageAsync(history, cts.Token);
            });

        // Show debug log if there are entries
        if (agent.DebugLog.Count > 0)
        {
            var debugTable = new Table()
                .BorderColor(Color.Grey)
                .AddColumn(new TableColumn("[dim]Debug Log[/]").NoWrap());
            
            debugTable.Border = TableBorder.Rounded;
            
            foreach (var logEntry in agent.DebugLog)
            {
                debugTable.AddRow(new Markup($"[dim]{logEntry.EscapeMarkup()}[/]"));
            }
            
            AnsiConsole.Write(debugTable);
            AnsiConsole.WriteLine();
        }

        var lastMessage = history.Last();
        if (lastMessage.Role == AuthorRole.Assistant)
        {
            var responsePanel = new Panel(new Markup($"[blue]{lastMessage.Content.EscapeMarkup()}[/]"))
                .Header("[blue bold]Response[/]", Justify.Left)
                .Border(BoxBorder.Rounded)
                .BorderColor(Color.Blue)
                .Padding(1, 0);
            
            AnsiConsole.Write(responsePanel);
            AnsiConsole.WriteLine();
        }
    }
    catch (Exception ex)
    {
        var errorPanel = new Panel($"[red bold]Error:[/] [red]{ex.Message.EscapeMarkup()}[/]\n[yellow]Tip: Make sure your LLM server is running and accessible.[/]")
            .Border(BoxBorder.Rounded)
            .BorderColor(Color.Red)
            .Padding(1, 0);
        
        AnsiConsole.Write(errorPanel);
        AnsiConsole.WriteLine();
        
        // Remove the failed user message from history
        if (history.Count > 0 && history.Last().Role == AuthorRole.User)
        {
            history.RemoveAt(history.Count - 1);
        }
    }
}

AnsiConsole.MarkupLine("[cyan]Session ended.[/]");
return 0;

static void RenderHeader(string workbookPath, string configuredModel, string actualModel)
{
    // Fixed width for modern terminals (100 chars)
    const int bannerWidth = 100;
    
    var grid = new Grid();
    grid.AddColumn(new GridColumn().Width(bannerWidth));
    
    // Chunkier title using block characters - centered for wider display
    var titlePadding = (bannerWidth - 72) / 2; // 72 is approx width of "WORKBOOK" line
    var pad = new string(' ', Math.Max(0, titlePadding));
    
    grid.AddRow(new Markup(""));
    grid.AddRow(new Markup($"{pad}[green1 bold]██╗      ██████╗  ██████╗ █████╗ ██╗         [/]"));
    grid.AddRow(new Markup($"{pad}[green1 bold]██║     ██╔═══██╗██╔════╝██╔══██╗██║         [/]"));
    grid.AddRow(new Markup($"{pad}[green1 bold]██║     ██║   ██║██║     ███████║██║         [/]"));
    grid.AddRow(new Markup($"{pad}[green1 bold]██║     ██║   ██║██║     ██╔══██║██║         [/]"));
    grid.AddRow(new Markup($"{pad}[green1 bold]███████╗╚██████╔╝╚██████╗██║  ██║███████╗    [/]"));
    grid.AddRow(new Markup($"{pad}[green1 bold]╚══════╝ ╚═════╝  ╚═════╝╚═╝  ╚═╝╚══════╝    [/]"));
    
    grid.AddRow(new Markup($"{pad}[green3 bold]██╗    ██╗ ██████╗ ██████╗ ██╗  ██╗██████╗  ██████╗  ██████╗ ██╗  ██╗[/]"));
    grid.AddRow(new Markup($"{pad}[green3 bold]██║    ██║██╔═══██╗██╔══██╗██║ ██╔╝██╔══██╗██╔═══██╗██╔═══██╗██║ ██╔╝[/]"));
    grid.AddRow(new Markup($"{pad}[green3 bold]██║ █╗ ██║██║   ██║██████╔╝█████╔╝ ██████╔╝██║   ██║██║   ██║█████╔╝ [/]"));
    grid.AddRow(new Markup($"{pad}[green3 bold]██║███╗██║██║   ██║██╔══██╗██╔═██╗ ██╔══██╗██║   ██║██║   ██║██╔═██╗ [/]"));
    grid.AddRow(new Markup($"{pad}[green3 bold]╚███╔███╔╝╚██████╔╝██║  ██║██║  ██╗██████╔╝╚██████╔╝╚██████╔╝██║  ██╗[/]"));
    grid.AddRow(new Markup($"{pad}[green3 bold] ╚══╝╚══╝  ╚═════╝ ╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝  ╚═════╝  ╚═════╝ ╚═╝  ╚═╝[/]"));
    
    grid.AddRow(new Markup($"{pad}[chartreuse1 bold] ██████╗██╗  ██╗ █████╗ ████████╗[/]"));
    grid.AddRow(new Markup($"{pad}[chartreuse1 bold]██╔════╝██║  ██║██╔══██╗╚══██╔══╝[/]"));
    grid.AddRow(new Markup($"{pad}[chartreuse1 bold]██║     ███████║███████║   ██║   [/]"));
    grid.AddRow(new Markup($"{pad}[chartreuse1 bold]██║     ██╔══██║██╔══██║   ██║   [/]"));
    grid.AddRow(new Markup($"{pad}[chartreuse1 bold]╚██████╗██║  ██║██║  ██║   ██║   [/]"));
    grid.AddRow(new Markup($"{pad}[chartreuse1 bold] ╚═════╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝   [/]"));
    
    grid.AddRow(new Markup(""));
    grid.AddRow(new Markup("[dim italic]AI-powered spreadsheet analysis - Private, Fast, Terminal-based[/]"));
    grid.AddRow(new Markup(""));
    
    // Session info table
    var infoTable = new Table()
        .HideHeaders()
        .AddColumn(new TableColumn("").Width(15))
        .AddColumn(new TableColumn(""));
    
    infoTable.Border = TableBorder.None;
    
    infoTable.AddRow(
        new Markup("[dim]Workbook:[/]"),
        new Markup($"[yellow bold]{Path.GetFileName(workbookPath)}[/] [dim]({new FileInfo(workbookPath).Length / 1024}KB)[/]")
    );
    
    // Show actual running model or warning
    if (!string.IsNullOrEmpty(actualModel) && actualModel != "unknown")
    {
        infoTable.AddRow(
            new Markup("[dim]AI Model:[/]"),
            new Markup($"[cyan bold]{actualModel}[/]")
        );
    }
    else
    {
        infoTable.AddRow(
            new Markup("[dim]AI Model:[/]"),
            new Markup($"[red bold]⚠ Please start LLM server on port 1234[/]")
        );
    }
    
    infoTable.AddRow(
        new Markup("[dim]Commands:[/]"),
        new Markup("[green]help[/] [dim]|[/] [green]load[/] [dim]<path>[/] [dim]|[/] [green]clear[/] [dim]|[/] [green]exit[/]")
    );
    
    grid.AddRow(infoTable);
    
    // Fixed-width panel (80 chars - Panel doesn't support .Width(), so use grid constraint)
    var panel = new Panel(grid)
        .Border(BoxBorder.Double)
        .BorderColor(Color.Green)
        .Padding(1, 0);
    
    AnsiConsole.Write(panel);
    AnsiConsole.WriteLine();
}

static async Task<string> DetectRunningModel(string baseUrl)
{
    try
    {
        using var httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(2) };
        var modelsUrl = baseUrl.Replace("/v1", "") + "/v1/models";
        
        var response = await httpClient.GetStringAsync(modelsUrl);
        
        // Parse JSON to get model name
        if (response.Contains("\"id\""))
        {
            var idStart = response.IndexOf("\"id\"") + 6;
            var idEnd = response.IndexOf("\"", idStart);
            if (idEnd > idStart)
            {
                return response.Substring(idStart, idEnd - idStart);
            }
        }
        
        return "unknown";
    }
    catch
    {
        return "unknown";
    }
}

static void ShowHelp()
{
    var table = new Table()
        .Border(TableBorder.Rounded)
        .BorderColor(Color.Yellow)
        .AddColumn("[green]Command[/]")
        .AddColumn("[green]Description[/]");

    table.AddRow("help, ?", "Show this help message");
    table.AddRow("clear, cls", "Clear the screen and redraw header");
    table.AddRow("load <path>", "Switch to a different workbook");
    table.AddRow("open <path>", "Alias for 'load' command");
    table.AddRow("switch", "Prompt for a new workbook path");
    table.AddRow("exit, quit, q", "Exit the application");
    table.AddRow("[dim]<message>[/]", "Ask a question about the workbook");

    AnsiConsole.Write(table);
    AnsiConsole.WriteLine();

    AnsiConsole.MarkupLine("[cyan]Example queries:[/]");
    AnsiConsole.MarkupLine("  • What sheets are in this workbook?");
    AnsiConsole.MarkupLine("  • Show me the first 10 rows of the Sales table");
    AnsiConsole.MarkupLine("  • Search for 'laptop' across all sheets");
    AnsiConsole.WriteLine();
    
    AnsiConsole.MarkupLine("[cyan]Workbook switching:[/]");
    AnsiConsole.MarkupLine("  • [green]load[/] D:/Projects/ProjectTracking.xlsx");
    AnsiConsole.MarkupLine("  • [green]open[/] \"C:/Sales Data/Q4-2024.xlsx\"");
    AnsiConsole.MarkupLine("  • [green]switch[/] (will prompt for path)");
    AnsiConsole.WriteLine();
}

static string? ResolveWorkbookPath(string[] arguments)
{
    if (arguments is null || arguments.Length == 0)
    {
        return Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
    }

    for (var i = 0; i < arguments.Length; i++)
    {
        var arg = arguments[i];
        if (string.Equals(arg, "--workbook", StringComparison.OrdinalIgnoreCase) || 
            string.Equals(arg, "-w", StringComparison.OrdinalIgnoreCase))
        {
            if (i + 1 < arguments.Length)
            {
                return arguments[i + 1];
            }
            break;
        }

        if (arg.StartsWith("--workbook=", StringComparison.OrdinalIgnoreCase))
        {
            return arg["--workbook=".Length..];
        }
    }

    return Environment.GetEnvironmentVariable("EXCEL_MCP_WORKBOOK");
}

static AgentConfiguration? LoadConfiguration()
{
    var baseUrl = Environment.GetEnvironmentVariable("LLM_BASE_URL");
    var modelId = Environment.GetEnvironmentVariable("LLM_MODEL_ID");
    var apiKey = Environment.GetEnvironmentVariable("LLM_API_KEY");

    if (!string.IsNullOrWhiteSpace(baseUrl) && !string.IsNullOrWhiteSpace(modelId))
    {
        return new AgentConfiguration
        {
            BaseUrl = baseUrl,
            ModelId = modelId,
            ApiKey = apiKey ?? "not-used"
        };
    }

    return null;
}

