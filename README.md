# NOTE: Changes still in progress
May cause breaking changes

# Excel Local MCP

Local Model Context Protocol server and CLI tools for conversational analysis of Excel workbooks. Chat with your spreadsheets using natural language - ask questions, search data, preview tables, and switch between workbooks without leaving the terminal.

**🚀 Current Status:** Production-ready CLI tools with debug logging, workbook switching, and AS/400-style terminal interface. Web UI is experimental.

## Features

✅ **Semantic Kernel CLI Agent** - AS/400-style terminal for conversational workbook queries  
✅ **Debug Logging** - See exactly which tools the LLM calls and when  
✅ **Workbook Switching** - Load different workbooks without restarting  
✅ **MCP Server** - Standards-compliant Model Context Protocol server  
✅ **Sample Workbooks** - Realistic test data (projects, employees, budgets)  
✅ **Local-First** - Everything runs on your machine, data never leaves  
✅ **Cross-Platform** - Windows, Linux, macOS via .NET 9  

> **Quick Start:** Jump to [Getting Started](#getting-started) to run in 3 minutes.

---

## Why This Exists

- **Keep spreadsheets private.** Many teams cannot upload financial or regulated Excel files to hosted services. Running the MCP server locally lets agents analyze data without leaving the device.
- **Bring Excel into the MCP ecosystem.** Most existing MCP tools focus on text documents or REST APIs. This project fills the gap by translating workbook structure into MCP tools/resources that any compliant client can consume.
- **Enable agent automation.** With a consistent schema for worksheets, tables, and rows, agents can answer natural language questions, generate summaries, and trigger downstream workflows that depend on spreadsheet context.

## Use Cases

- “What tabs and tables exist in the budget workbook, and who owns each one?”
- “Find every row across worksheets that references supplier X with an overdue balance.”
- “Preview the first 20 rows of the quarterly awards table and send it to a teammate.”
- “Combine this with a local CRM MCP server so the agent can reconcile spreadsheet exports with live system data.”

## Future Enhancements

See the evolving roadmap in [docs/FutureFeatures.md](docs/FutureFeatures.md). Highlights on deck:

- Support filtered range previews (e.g., formulas vs. values, pivot expansions).
- Implement write-back tools to update cells, add worksheets, or annotate findings.
- Expose analytics such as value distributions, outlier detection, or chart generation.
- Add WebSocket/HTTP transports so the server can run behind a relay or container orchestration platform.

## Components

### Production-Ready
- **`src/ExcelMcp.Server`** – Stdio JSON-RPC MCP server that indexes a workbook
- **`src/ExcelMcp.Client`** – Command-line tool that starts the server and calls MCP tools/resources
- **`src/ExcelMcp.SkAgent`** – Semantic Kernel CLI agent with AS/400-style terminal interface for conversational workbook analysis (includes debug logging, workbook switching, and colorized output)
- **`src/ExcelMcp.Contracts`** – Shared data contracts

### Documentation
- **`docs/UserGuide.md`** – Extended walkthrough covering setup, workflows, and troubleshooting
- **`docs/FutureFeatures.md`** – Forward-looking ideas we plan to explore
- **`docs/SkAgentQuickStart.md`** – Quick start guide for the Semantic Kernel agent
- **`docs/SkAgentDebugLog.md`** – Debug logging feature documentation
- **`test-data/README.md`** – Sample workbooks and testing guide

### Under Development
- **`src/ExcelMcp.ChatWeb`** – ASP.NET web UI for chat interface (experimental, not feature-complete)

## Prerequisites

- .NET SDK 9.0+
- PowerShell 7+ (Windows/macOS/Linux) for the packaging scripts
- Local `.xlsx` workbook to analyze

## Build & Test

```pwsh
dotnet build
dotnet test
```

## Quick Starts

### Semantic Kernel CLI Agent (Recommended)

The SK agent provides a classic AS/400-style terminal interface with conversational AI, debug logging, and workbook switching:

```pwsh
# 1. Start your local LLM server (e.g., LM Studio on port 1234)

# 2. Create sample workbooks (optional)
pwsh -File scripts/create-sample-workbooks.ps1

# 3. Run the agent
dotnet run --project src/ExcelMcp.SkAgent -- --workbook "test-data/ProjectTracking.xlsx"

# 4. Chat with your workbook:
> what tables exist in this workbook?
> show me all high priority tasks
> load test-data/EmployeeDirectory.xlsx
> who works in Engineering?
> exit
```

**Features:**
- ✅ Green colorized terminal UI
- ✅ Debug logging shows tool calls
- ✅ Switch workbooks without restarting (`load`, `open`, `switch` commands)
- ✅ Case-insensitive search
- ✅ Works with local LLMs (LM Studio, Ollama) or OpenAI

**See:** [docs/SkAgentQuickStart.md](docs/SkAgentQuickStart.md) for detailed guide.

### MCP Server and CLI Client

For direct MCP protocol interaction or programmatic access:

```pwsh
# Use the CLI client (prompts for workbook path on first run)
dotnet run --project src/ExcelMcp.Client -- list
dotnet run --project src/ExcelMcp.Client -- search "product name"
dotnet run --project src/ExcelMcp.Client -- preview Sales --rows 10
```

### Running the MCP Server Directly

For integration with other MCP-compatible clients:

```pwsh
# Run the stdio MCP server directly
dotnet run --project src/ExcelMcp.Server -- --workbook "D:/Data/sample.xlsx"
```

## Example mcp,json for VS Code

```
{
  "servers": {
    "excel-workbook-mcp": {
      "type": "stdio",
      "command": "${workspaceFolder}/src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server.exe",
      "args": [
        "--workbook",
        "${input:excel-workbook-path}"
      ]
    }
  },
  "inputs": [
    {
      "id": "excel-workbook-path",
      "type": "promptString",
      "description": "Full path to the Excel workbook to load",
      "default": "D:/Downloads/sampledata.xlsx"
    }
  ]
}
```

## Packaging

Package each component as a self-contained, single-file executable:

```pwsh
pwsh -File scripts/package-server.ps1   # MCP Server
pwsh -File scripts/package-client.ps1   # CLI Client
pwsh -File scripts/package-skagent.ps1  # Semantic Kernel Agent (Recommended!)
```

Each script publishes to `dist/<rid>/<AppName>` with platform-specific launch helpers (`.ps1`, `.sh`, `.bat`).

**Note:** The web chat (`package-chatweb.ps1`) is experimental and not feature-complete.
pwsh -File scripts/package-chatweb.ps1  # ExcelMcp.ChatWeb (includes wwwroot)
pwsh -File scripts/package-skagent.ps1  # ExcelMcp.SkAgent (Semantic Kernel terminal agent)
```

Pass `-Runtime` (e.g., `linux-x64`) or `-SkipZip` as needed. The generated folders include `run-*.ps1`, `run-*.sh`, and `run-*.bat` wrappers that prompt for workbook/server paths and start the bundled executable.

## Integrate with Agents

Point your MCP-capable agent at the packaged server or use the CLI/web app launchers to negotiate workbook/server paths. All tools (`excel-list-structure`, `excel-search`, `excel-preview-table`) and resources (`excel://` URIs) follow the MCP spec, so they work with OpenAI agents, MCP bridges, or any compatible orchestrator.

## Getting Started

### 1. Build the Project
```pwsh
dotnet build
dotnet test
```

### 2. Create Sample Workbooks
```pwsh
pwsh -File scripts/create-sample-workbooks.ps1
```

This creates three test workbooks in `test-data/`:
- **ProjectTracking.xlsx** - Tasks, projects, time logs
- **EmployeeDirectory.xlsx** - Employee and department data
- **BudgetTracker.xlsx** - Income and expense tracking

See [test-data/README.md](test-data/README.md) for details and example queries.

### 3. Run the Semantic Kernel Agent
```pwsh
# With a local LLM (recommended: LM Studio, Ollama)
dotnet run --project src/ExcelMcp.SkAgent -- --workbook test-data/ProjectTracking.xlsx

# Or set environment variables
$env:LLM_BASE_URL = "http://localhost:1234/v1"
$env:LLM_MODEL_ID = "local-model"
$env:EXCEL_MCP_WORKBOOK = "test-data/ProjectTracking.xlsx"
dotnet run --project src/ExcelMcp.SkAgent
```

### 4. Try It Out
```
> what tasks are high priority?
> show me all employees in Engineering
> load test-data/BudgetTracker.xlsx
> what's the total income?
> help
```

**Key Features to Explore:**
- 🔧 **Debug logging** - See exactly which tools the LLM calls
- 🔄 **Workbook switching** - Use `load`, `open`, or `switch` commands
- 🎨 **Colorized output** - Green banner, color-coded messages
- 📊 **Sample workbooks** - Realistic test data included

## Documentation

- **[User Guide](docs/UserGuide.md)** - Complete setup and workflow guide
- **[SK Agent Quick Start](docs/SkAgentQuickStart.md)** - CLI agent tutorial
- **[Debug Logging](docs/SkAgentDebugLog.md)** - Understanding tool calls
- **[Sample Workbooks](test-data/README.md)** - Test data reference
- **[Future Features](docs/FutureFeatures.md)** - Planned enhancements

## Troubleshooting

Common issues and solutions:

**LLM not calling tools / making up data:**
- See [docs/SkAgentTroubleshooting.md](docs/SkAgentTroubleshooting.md)
- Try better models: `gpt-4`, `phi-4`, `llama-3.1-8b-instruct`
- Check debug log: `⚠️ No tools were called` means model issue

**Search returning no results:**
- Fixed! Empty string bug resolved
- Search is case-insensitive by default
- Debug log shows: `caseSensitive=False`

**Wrong calculations:**
- Model limitation - local models may struggle with math
- Debug log shows tool calls but incorrect computation
- Use GPT-4 for accurate calculations

Refer to [docs/UserGuide.md](docs/UserGuide.md) for detailed workflow examples, environment variables, and troubleshooting notes.

## Acknowledgements

This project relies on the excellent ClosedXML library (https://github.com/ClosedXML/ClosedXML) as its core engine for reading and writing Excel workbooks (.xlsx files).
ClosedXML provides a clean, intuitive API built on top of the OpenXML SDK, allowing this MCP server to interact with Excel data without requiring Microsoft Excel or any COM automation.

Every workbook operation exposed through this MCP server, such as reading rows, writing cell values, or adding data, is powered internally by ClosedXML.

A huge thank you to the ClosedXML maintainers and contributors for their ongoing work on one of the most reliable and developer-friendly Excel libraries in the .NET ecosystem.
Your project makes it possible for tools like this to exist and run cross-platform in lightweight environments