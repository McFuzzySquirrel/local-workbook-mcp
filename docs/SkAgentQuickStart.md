# Excel MCP Semantic Kernel Agent - Quick Start Guide

**Last Updated:** March 12, 2026

This guide shows you how to get started with the AS/400-style terminal agent powered by Semantic Kernel.

## What You'll Need

1. **A local LLM server** — [Ollama](https://ollama.com/) (recommended) or [LM Studio](https://lmstudio.ai/)
   - Ollama default: `http://localhost:11434`
   - LM Studio default: `http://localhost:1234`
   - Load a model with good function-calling support (see recommendations below)

2. **An Excel workbook** to analyze (`.xlsx` format)

3. **.NET 10.0 SDK** installed

## Quick Start

### Step 1: Start Your LLM Server

**Option A: Ollama (Recommended)**
```bash
ollama pull llama3.2
ollama serve
```

**Option B: LM Studio**
1. Open LM Studio and download a model (e.g., `phi-4`, `llama-3.2-3b-instruct`)
2. Start the local server (default port 1234)
3. Verify it's running on `http://localhost:1234`

### Step 2: Run the Agent

```bash
# From the repository root
dotnet run --project src/ExcelMcp.SkAgent -- --workbook "path/to/your/workbook.xlsx"

# Or use environment variable
export EXCEL_MCP_WORKBOOK="/home/user/data/sample.xlsx"
dotnet run --project src/ExcelMcp.SkAgent
```

### Step 3: Chat with Your Workbook

Once the agent starts, you'll see a green ASCII art banner and a prompt (`>`). Try these queries:

```
> What sheets are in this workbook?
> Show me the first 10 rows of the Sales table
> Search for laptop
> How many tables are in this workbook?
> Update cell A1 in Sheet1 to "Updated Value"
> Create a new worksheet called Summary
> Analyze the pivot table in SalesPivot
> exit
```

## Terminal Commands

- **help** or **?** - Show available commands
- **clear** or **cls** - Clear the screen
- **exit**, **quit**, or **q** - Exit the agent

## Configuration (Optional)

The agent auto-detects Ollama at `http://localhost:11434`. If you're using a different endpoint, set these environment variables:

```bash
# For a custom Ollama model
export LLM_BASE_URL="http://localhost:11434/v1"
export LLM_MODEL_ID="llama3.2"
export LLM_API_KEY=""   # not needed for local

# For LM Studio
export LLM_BASE_URL="http://localhost:1234/v1"
export LLM_MODEL_ID="local-model"
export LLM_API_KEY="lm-studio"

# PowerShell equivalents
$env:LLM_BASE_URL = "http://localhost:1234/v1"
$env:LLM_MODEL_ID = "local-model"
$env:LLM_API_KEY = "lm-studio"
```

## Recommended Models

For best function-calling results (in order of recommendation):

1. `llama3.2` or `llama3.1` via Ollama
2. `phi-4` via LM Studio or Ollama
3. `mistral` (Ollama)
4. `gpt-4` or `gpt-4-turbo` (OpenAI API — requires `LLM_API_KEY`)

## The AS/400 Experience

The agent is designed to feel like classic terminal computing:

- **Green ASCII art** banner on startup
- **Simple prompt** (`>`) for input
- **Command-based** interface (help, clear, exit)
- **Fast and lightweight** — no web browser needed
- **Text-based tables** for data display
- **Status indicators** with spinners during processing

## Behind the Scenes

When you ask a question, the agent uses these SK plugin functions:

| Function | Purpose |
|---|---|
| `list_structure` | Describe worksheets, tables, columns, pivot tables |
| `search` | Find text across the workbook |
| `preview_table` | Show rows from a worksheet or named table |
| `get_workbook_summary` | High-level workbook overview |
| `write_cell` | Update a single cell value |
| `write_range` | Update multiple cells in one operation |
| `create_worksheet` | Add a new blank worksheet |
| `analyze_pivot` | Inspect pivot table structure and data |

Each SK function wraps the corresponding MCP tool call. Write operations automatically create a timestamped `.xlsx` backup before saving.

## Example Session

```
  ___  __  __  ___  ___  _       __  __   ___  ___
 | __| \ \/ / / __|| __|| |     |  \/  | / __|| _ \
 | _|   >  < | (__ | _| | |__   | |\/| || (__ |  _/
 |___| /_/\_\ \___||___||____|  |_|  |_| \___||_|

Semantic Kernel Agent for Local Excel Workbooks
Type 'help' for commands, 'exit' to quit

Using workbook: /home/user/data/ProjectTracking.xlsx

> What sheets are in this workbook?
  The workbook contains 3 worksheets: Projects, Tasks, Resources

> Show me the first 5 rows of Projects
  | ID | Name           | Status  | Owner   |
  |----|----------------|---------|---------|
  | 1  | Alpha Launch   | Active  | Alice   |
  | 2  | Beta Migration | On Hold | Bob     |
  | 3  | Gamma Rollout  | Active  | Carol   |
  | 4  | Delta Upgrade  | Done    | Dave    |
  | 5  | Epsilon Prep   | Active  | Eve     |

> Update cell B2 in Projects to "Project Renew"
  ✓ Cell B2 updated. Backup created: ProjectTracking_2026-03-12T143025Z.xlsx

> exit
```
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
