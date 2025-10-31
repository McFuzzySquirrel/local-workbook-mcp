# Excel MCP Semantic Kernel Agent

A classic terminal-style CLI agent for conversing with Excel workbooks using Semantic Kernel and local LLMs.

## Features

- **AS/400-inspired terminal interface** - Green text, command-line feel, classic computing aesthetic
- **Semantic Kernel integration** - Automatic function calling and agent orchestration
- **Local LLM support** - Works with LM Studio, Ollama, or any OpenAI-compatible endpoint
- **Interactive REPL** - Conversational loop with chat history
- **Rich terminal UI** - Powered by Spectre.Console for beautiful tables and formatting
- **Zero web dependencies** - Runs anywhere .NET 9 runs

## Quick Start

```pwsh
# 1. Start your local LLM (e.g., LM Studio on http://localhost:1234)

# 2. Run the agent
dotnet run --project src/ExcelMcp.SkAgent -- --workbook "path/to/your/workbook.xlsx"

# Or let it prompt you for the workbook path
dotnet run --project src/ExcelMcp.SkAgent
```

## Configuration

The agent can be configured via environment variables:

```pwsh
# LLM endpoint configuration
$env:LLM_BASE_URL = "http://localhost:1234"
$env:LLM_MODEL_ID = "phi-4-mini-reasoning"
$env:LLM_API_KEY = "not-used"  # For local models

# Workbook path (optional)
$env:EXCEL_MCP_WORKBOOK = "D:/Data/sample.xlsx"
```

If not configured, defaults to LM Studio on localhost:1234.

## Commands

While in the agent REPL:

- **help, ?** - Show available commands
- **clear, cls** - Clear the screen
- **exit, quit, q** - Exit the application
- **Any other text** - Ask the agent about the workbook

## Example Queries

```
> What sheets are in this workbook?
> Show me the first 10 rows of the Sales table
> Search for 'laptop' across all sheets
> How many rows are in the Products table?
> What columns does the Customers table have?
```

## How It Works

1. The agent loads the Excel workbook using `ExcelWorkbookService`
2. Semantic Kernel registers four functions as tools:
   - `list_structure` - List all worksheets and tables
   - `search` - Search for text across the workbook
   - `preview_table` - Preview rows from a table/worksheet
   - `get_workbook_summary` - Get metadata about the workbook
3. Your LLM receives these tool descriptions and automatically calls them
4. Results are formatted and displayed in the terminal

## Building

```pwsh
dotnet build src/ExcelMcp.SkAgent/ExcelMcp.SkAgent.csproj
```

## Packaging

```pwsh
pwsh -File scripts/package-skagent.ps1
```

This creates a self-contained executable in `dist/win-x64/ExcelMcp.SkAgent/` (or your target runtime).

## Requirements

- .NET 9.0+
- Local LLM server (LM Studio, Ollama, etc.) or OpenAI API key
- Excel workbook (.xlsx format)

**Important:** Use LM Studio 0.3.0+ for best compatibility. See [Troubleshooting Guide](../../docs/SkAgentTroubleshooting.md) if you encounter issues.

## Why Terminal UI?

This agent is designed for:
- **Portability** - Runs on servers, containers, SSH sessions
- **Simplicity** - No browser, no web server, just a terminal
- **Classic feel** - AS/400-style green screen aesthetic for that retro computing vibe
- **Performance** - Lightweight, fast startup, minimal resource usage
- **Scriptability** - Can be piped, redirected, and integrated into workflows
