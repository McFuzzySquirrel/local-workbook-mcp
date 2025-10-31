# Excel MCP Semantic Kernel Agent - Quick Demo Guide

This guide shows you how to get started with the new AS/400-style terminal agent.

## What You'll Need

1. **A local LLM server** running on `http://localhost:1234`
   - [LM Studio](https://lmstudio.ai/) is recommended (free, easy to use)
   - Load a model like `phi-4-mini-reasoning` or any other chat model
   
2. **An Excel workbook** to analyze (`.xlsx` format)

3. **.NET 9.0 SDK** installed

## Quick Start

### Step 1: Start Your LLM Server

```pwsh
# If using LM Studio:
# 1. Open LM Studio
# 2. Download a model (e.g., phi-4-mini-reasoning)
# 3. Start the local server (default port 1234)
# 4. Make sure the server is running and ready
```

### Step 2: Run the Agent

```pwsh
# From the repository root
dotnet run --project src/ExcelMcp.SkAgent -- --workbook "path/to/your/workbook.xlsx"

# Or use environment variable
$env:EXCEL_MCP_WORKBOOK = "D:/Data/sample.xlsx"
dotnet run --project src/ExcelMcp.SkAgent
```

### Step 3: Chat with Your Workbook

Once the agent starts, you'll see a green ASCII art banner and a prompt (`>`). Try these queries:

```
> What sheets are in this workbook?
> Show me the first 10 rows of the Sales table
> Search for laptop
> How many tables are in this workbook?
> exit
```

## Terminal Commands

- **help** or **?** - Show available commands
- **clear** or **cls** - Clear the screen
- **exit**, **quit**, or **q** - Exit the agent

## Configuration (Optional)

If you're using a different LLM endpoint, set these environment variables:

```pwsh
$env:LLM_BASE_URL = "http://localhost:11434"  # e.g., for Ollama
$env:LLM_MODEL_ID = "llama2"
$env:LLM_API_KEY = "your-api-key-if-needed"
```

## The AS/400 Experience

The agent is designed to feel like classic terminal computing:

- **Green ASCII art** banner on startup
- **Simple prompt** (`>`) for input
- **Command-based** interface (help, clear, exit)
- **Fast and lightweight** - no web browser needed
- **Text-based tables** for data display
- **Status indicators** with spinners during processing

## Behind the Scenes

When you ask a question:

1. **Semantic Kernel** receives your query
2. The **LLM decides** which Excel tools to call
3. **Tools are invoked** automatically:
   - `list_structure` - Get worksheet/table metadata
   - `search` - Find text across the workbook
   - `preview_table` - Show rows from tables
   - `get_workbook_summary` - Get workbook overview
4. **Results formatted** and displayed in the terminal

## Example Session

```
  ___  __  __  ___  ___  _       __  __   ___  ___
 | __| \ \/ / / __|| __|| |     |  \/  | / __|| _ \
 | _|   >  < | (__ | _| | |__   | |\/| || (__ |  _/
 |___| /_/\_\ \___||___||____|  |_|  |_| \___||_|

Semantic Kernel Agent for Local Excel Workbooks
Type 'help' for commands, 'exit' to quit

═══════════════════════════════════════════════════════════
Workbook: sample-sales-data.xlsx
Model: phi-4-mini-reasoning
═══════════════════════════════════════════════════════════

> What sheets are in this workbook?

The workbook contains 3 sheets:
1. Sales - Contains sales transaction data
2. Products - Contains product information  
3. Customers - Contains customer records

> Show me the first 5 rows of the Sales table

Here are the first 5 rows of the Sales table:

OrderID | Date       | Product    | Quantity | Price
--------|------------|------------|----------|-------
1001    | 2024-01-15 | Laptop     | 2        | $999
1002    | 2024-01-16 | Mouse      | 10       | $29
1003    | 2024-01-16 | Keyboard   | 5        | $79
1004    | 2024-01-17 | Monitor    | 3        | $299
1005    | 2024-01-18 | Laptop     | 1        | $999

> exit

Session ended.
```

## Packaging for Distribution

To create a standalone executable:

```pwsh
# Package for Windows
pwsh -File scripts/package-skagent.ps1 -Runtime win-x64

# Package for Linux
pwsh -File scripts/package-skagent.ps1 -Runtime linux-x64

# Package for macOS
pwsh -File scripts/package-skagent.ps1 -Runtime osx-x64
```

The packaged agent will be in `dist/<runtime>/ExcelMcp.SkAgent/` with helper scripts.

## Troubleshooting

**"Connection refused" or similar errors:**
- Make sure your LLM server is running
- Check the URL matches (default: `http://localhost:1234`)
- Verify the model is loaded in your LLM server

**"Workbook not found":**
- Use full paths to workbooks
- Make sure the file is `.xlsx` format
- Check file permissions

**Tools not being called:**
- Some models work better than others for function calling
- Try `phi-4-mini-reasoning` or `gpt-4` for best results
- Check LLM server logs for errors

## Why Terminal UI?

This agent gives you:
- ✅ **Privacy** - Everything runs locally
- ✅ **Portability** - Works on servers, SSH, containers
- ✅ **Speed** - Fast startup, low memory
- ✅ **Simplicity** - No browser, no web server
- ✅ **Classic feel** - Retro computing aesthetic
- ✅ **Scriptability** - Can be automated and piped

Perfect for data analysts, system administrators, and anyone who loves working in the terminal!
