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
RenderHeader(workbookPath, config.ModelId);

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
        RenderHeader(workbookPath, config.ModelId);
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
            RenderHeader(workbookPath, config.ModelId);
            
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

static void RenderHeader(string workbookPath, string modelId)
{
    var grid = new Grid();
    grid.AddColumn();
    
    // Colored ASCII art banner - GitHub Copilot style
    grid.AddRow(new Markup(""));
    grid.AddRow(new Markup("[green1 bold] ╦   ┌─┐┌─┐┌─┐┬    [/][green3 bold] ╦ ╦┌─┐┬─┐┬┌─┌┐ ┌─┐┌─┐┬┌─ [/][chartreuse1 bold] ╔═╗┬ ┬┌─┐┌┬┐[/]"));
    grid.AddRow(new Markup("[green1 bold] ║   │ ││  ├─┤│    [/][green3 bold] ║║║│ │├┬┘├┴┐├┴┐│ ││ │├┴┐ [/][chartreuse1 bold] ║  ├─┤├─┤ │ [/]"));
    grid.AddRow(new Markup("[green1 bold] ╩═╝ └─┘└─┘┴ ┴┴─┘  [/][green3 bold] ╚╩╝└─┘┴└─┴ ┴└─┘└─┘└─┘┴ ┴ [/][chartreuse1 bold] ╚═╝┴ ┴┴ ┴ ┴ [/]"));
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
    
    infoTable.AddRow(
        new Markup("[dim]AI Model:[/]"),
        new Markup($"[cyan bold]{modelId}[/]")
    );
    
    infoTable.AddRow(
        new Markup("[dim]Commands:[/]"),
        new Markup("[green]help[/] [dim]|[/] [green]load[/] [dim]<path>[/] [dim]|[/] [green]clear[/] [dim]|[/] [green]exit[/]")
    );
    
    grid.AddRow(infoTable);
    
    // Auto-expand panel to terminal width
    var panel = new Panel(grid)
        .Border(BoxBorder.Double)
        .BorderColor(Color.Green)
        .Padding(1, 0)
        .Expand(); // This makes it expand to terminal width
    
    AnsiConsole.Write(panel);
    AnsiConsole.WriteLine();
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

