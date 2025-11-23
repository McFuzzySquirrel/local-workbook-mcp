# Excel Local MCP - User Guide

**Last Updated:** November 1, 2025

This guide covers setup, workflows, and troubleshooting for the Excel Local MCP project.

**Status:** 
- ✅ **CLI Agent** - Stable, well-tested, recommended for daily use
- ⚠️ **Web Chat** - Functional, needs validation testing before broad use

## Prerequisites

- .NET 9.0 SDK (for building from source) or self-contained executables from the distribution
- A local LLM server for the SK agent (e.g., [LM Studio](https://lmstudio.ai/), Ollama)
- Excel workbook files (`.xlsx` format)

## Getting Started with the Semantic Kernel CLI Agent (Recommended)

The SK agent provides an AS/400-style terminal interface for conversational workbook analysis.

### 1. Start Your Local LLM Server

The SK agent requires an OpenAI-compatible API endpoint. Recommended options:

**Option A: LM Studio (Easiest)**
1. Download and install [LM Studio](https://lmstudio.ai/)
2. Download a model (recommended: `phi-4`, `llama-3.1-8b-instruct`, or `gpt-4` compatible)
3. Load the model and start the local server
4. Verify it's running on `http://localhost:1234`

**Option B: Ollama**
```bash
ollama serve
ollama run llama3.2
```

**Option C: OpenAI API**
Set your API key in environment variables (see Configuration section below).

### 2. Configure the SK Agent (Optional)

The agent uses sensible defaults, but you can customize via environment variables:

### 3. Launch the Web UI

#### From Source:

```powershell
dotnet run --project src/ExcelMcp.ChatWeb
```

#### From Distribution:

```powershell
# Windows
cd dist/win-x64/ExcelMcp.ChatWeb
./run-chatweb.ps1

# Linux/macOS
cd dist/linux-x64/ExcelMcp.ChatWeb
chmod +x run-chatweb.sh
./run-chatweb.sh
```

### 4. Use the Web Interface

1. Open your browser to `http://localhost:5000`
2. Click **"Choose Excel file..."** in the right sidebar
3. Select your Excel workbook (`.xlsx` or `.xls`)
4. Click **"Load Workbook"**
5. Once loaded, you'll see:
   - Workbook name and sheet count
   - List of sheets in the workbook
   - Suggested questions you can ask

### 5. Ask Questions

Type natural language queries in the input box:

**Metadata queries:**
- "What sheets are in this workbook?"
- "How many tables are there?"
- "What columns are in the Sales sheet?"

**Data retrieval:**
- "Show me the first 10 rows of the Sales table"
- "Display all rows from the Inventory sheet"
- "Preview the Returns table"

**Search queries:**
- "Search for Laptop in the workbook"
- "Find all mentions of Contoso"
- "Where does 'revenue' appear?"

**Advanced Filtering:**
- "Show sales greater than 5000"
- "Find customers in NY"
- "Show items with price between 10 and 20"

**Cross-Sheet Analysis:**
- "Compare Sales and Inventory for matching products"
- "Find all mentions of Project X across all sheets"

### 6. Advanced Features

- **Summarize:** Click the "Summarize" button to get a concise report of your conversation and findings.
- **Export Chat:** Click "Export Chat" to download the full conversation history as a Markdown file.
- **Export Data:** When viewing a table, click "Export CSV" to download the data.
- **Context Awareness:** The agent remembers previous questions. You can ask "Show me the first one" after listing sheets.

The system will:
- Automatically select the right MCP tool based on your question
- Call the MCP server with appropriate parameters
- Format the response (tables are rendered as HTML)
- Maintain conversation history

## Using the CLI Client

For scripting or quick queries without the web UI:

```powershell
# List workbook structure
dotnet run --project src/ExcelMcp.Client -- list

# Search for text
dotnet run --project src/ExcelMcp.Client -- search "product name"

# Preview table rows
dotnet run --project src/ExcelMcp.Client -- preview Sales --rows 10
dotnet run --project src/ExcelMcp.Client -- preview Inventory --table InventoryTable

# Get resource content
dotnet run --project src/ExcelMcp.Client -- resources
dotnet run --project src/ExcelMcp.Client -- get excel://worksheet/Sales
```

On first run, the client will prompt for:
1. **Workbook path** - Full path to your Excel file
2. **Server path** - Path to `ExcelMcp.Server.exe` (auto-detected if running from source)

These values are cached in environment variables for subsequent runs.

## Running the MCP Server Standalone

The server can run independently for integration with other MCP-compatible tools:

```powershell
# Direct execution
dotnet run --project src/ExcelMcp.Server -- --workbook "D:/Data/finance.xlsx"

# Or from distribution
./ExcelMcp.Server.exe --workbook "D:/Data/finance.xlsx"
```

The server communicates via stdin/stdout using JSON-RPC 2.0. It exposes:

### MCP Tools

1. **excel-list-structure** - Returns JSON metadata about all worksheets, tables, columns
2. **excel-search** - Full-text search across all cells
3. **excel-preview-table** - Returns CSV data from a worksheet or named table

### MCP Resources

- `excel://workbook/metadata` - Workbook metadata
- `excel://worksheet/{name}` - Specific worksheet details

## Distribution and Packaging

### Build Self-Contained Executables

```powershell
# Package all components for Windows
pwsh -File scripts/package-server.ps1 -Runtime win-x64
pwsh -File scripts/package-client.ps1 -Runtime win-x64  
pwsh -File scripts/package-chatweb.ps1 -Runtime win-x64

# Package for Linux
pwsh -File scripts/package-server.ps1 -Runtime linux-x64
pwsh -File scripts/package-client.ps1 -Runtime linux-x64
pwsh -File scripts/package-chatweb.ps1 -Runtime linux-x64

# Package for multiple platforms at once
pwsh -File scripts/package-chatweb.ps1 -Runtime @('win-x64','linux-x64','osx-arm64')
```

Each script creates:
- Self-contained executable (no .NET runtime needed)
- Launch scripts (`.ps1`, `.bat`, `.sh`)
- README with usage instructions
- Optional zip archive (use `-SkipZip` to skip)

Outputs go to `dist/{runtime-id}/{AppName}/`

### Share with Users

Distribute the zip file or the entire `dist/{runtime-id}/{AppName}` folder. Users can run the launch scripts without installing .NET.

## Configuration Reference

### appsettings.json (ChatWeb)

```json
{
  "Logging": {
    "LogLevel": {
      "Default": "Information"
    }
  },
  "LlmStudio": {
    "BaseUrl": "http://localhost:1234/v1",
    "Model": "phi-4-mini-reasoning"
  },
  "ExcelMcp": {
    "ServerPath": "D:/path/to/ExcelMcp.Server.exe"
  }
}
```

### Environment Variables

- `EXCEL_MCP_WORKBOOK` - Default workbook path
- `EXCEL_MCP_SERVER` - Path to MCP server executable
- `LLM_STUDIO_BASE_URL` - Override LLM endpoint
- `LLM_STUDIO_MODEL` - Override LLM model name

## How It Works

### Architecture

```
┌─────────────────┐
│   Web Browser   │
│  (Blazor UI)    │
└────────┬────────┘
         │ HTTP/SignalR
         ↓
┌─────────────────────────┐
│  ExcelMcp.ChatWeb       │
│  - Blazor Server        │
│  - Semantic Kernel      │
│  - Agent Service        │
└────────┬────────────────┘
         │ stdio
         ↓
┌─────────────────────────┐      ┌──────────────┐
│  ExcelMcp.Server        │      │  LM Studio   │
│  - MCP Protocol         │      │  (OpenAI API)│
│  - ClosedXML            │      └──────────────┘
│  - JSON-RPC             │             ↑
└─────────────────────────┘             │
         │                               │
         ↓                               │
┌─────────────────────────┐             │
│   Excel Workbook        │             │
│   (.xlsx file)          │             │
└─────────────────────────┘             │
                                        │
         Semantic Kernel ───────────────┘
         connects ChatWeb to LLM
```

### Data Flow

1. **User uploads workbook** → ChatWeb saves to temp directory
2. **ChatWeb starts MCP server** → Passes workbook path via `--workbook` argument
3. **Server loads workbook** → Uses ClosedXML to parse Excel file
4. **User asks question** → ChatWeb sends to Semantic Kernel agent
5. **Agent analyzes query** → Determines which MCP tool to call
6. **Agent calls MCP tool** → ChatWeb forwards request to server via JSON-RPC
7. **Server returns data** → JSON metadata or CSV table data
8. **ChatWeb formats response** → Converts CSV to HTML tables, displays metadata
9. **User sees answer** → Rendered in chat interface with proper formatting

## Troubleshooting

Each component now has its own packaging script. From the repository root run whichever bundles you need, optionally passing one or more Runtime Identifiers (RIDs) such as `win-x64`, `linux-x64`, or `linux-arm64`.

```powershell
pwsh -File scripts/package-server.ps1
pwsh -File scripts/package-client.ps1 -Runtime @('win-x64','linux-x64')
pwsh -File scripts/package-chatweb.ps1 -Runtime linux-x64 -SkipZip
```

Each script publishes a self-contained single-file executable, copies launch helpers (`run-*.ps1`, `run-*.bat`, `run-*.sh`), writes a README, and optionally creates a zip archive.

Key outputs per RID:

- `dist/<rid>/ExcelMcp.Server/ExcelMcp.Server[.exe]`
- `dist/<rid>/ExcelMcp.Client/ExcelMcp.Client[.exe]`
- `dist/<rid>/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb[.exe]` plus `wwwroot/`

> **Tip:** The .NET runtime is bundled inside the published executable. Users do not need any additional installs.

## 2. Share with Users

Distribute either of the following:

- The zip file the script produced (`excel-mcp-<app>-<rid>.zip`), **or**
- The entire folder `dist/<rid>/<AppName>` copied to removable media or a shared drive.

Instruct users to extract or copy to a writable location, for example:

```
C:\Apps\ExcelMcp\win-x64
~/excel-mcp/linux-x64
```

Everything the app needs—including the ChatWeb `wwwroot` assets—is contained within that directory tree.

## 3. Use Cases and Examples

### 3.1 Launch the Chat Web UI (Windows)

```powershell
cd "C:\Apps\ExcelMcp\win-x64\ExcelMcp.ChatWeb"
./run-chatweb.ps1
```

- The launcher prompts for a workbook path (e.g., `D:\Data\finance.xlsx`).
- It locates the bundled server, sets environment variables, and starts `ExcelMcp.ChatWeb.exe`.
- Browse to the URL printed in the console (defaults to `http://localhost:5000`).

Double-click `run-chatweb.bat` for a shortcut that wraps the PowerShell script.

### 3.2 Launch the Chat Web UI (Linux / macOS)

```bash
cd ~/excel-mcp/linux-x64/ExcelMcp.ChatWeb
chmod +x run-chatweb.sh
./run-chatweb.sh --workbook /home/user/data/finance.xlsx --urls http://0.0.0.0:8080
```

The script exports `EXCEL_MCP_WORKBOOK`, points to the bundled server, and launches the ASP.NET app. Adjust `--urls` to bind to a different host/port.

### 3.3 Run the Server Manually (Any Platform)

Copy only the server folder onto a machine and start it with a workbook:

```powershell
cd "C:\Apps\ExcelMcp\win-x64\ExcelMcp.Server"
./ExcelMcp.Server.exe --workbook "D:\Data\finance.xlsx"
```

```bash
cd ~/excel-mcp/linux-x64/ExcelMcp.Server
./ExcelMcp.Server --workbook /home/user/data/finance.xlsx
```

- The process exposes MCP tools over standard input/output.
- If `--workbook` is omitted and the console is interactive, the executable prompts for a path.
- Stop with `Ctrl+C`.

### 3.4 Explore Tools with the CLI Client

```powershell
cd "C:\Apps\ExcelMcp\win-x64\ExcelMcp.Client"
./run-client.ps1 -WorkbookPath "D:\Data\finance.xlsx" list
```

```bash
cd ~/excel-mcp/linux-x64/ExcelMcp.Client
./run-client.sh --workbook /home/user/data/finance.xlsx search --query "Contoso"
```

The client wrapper prompts for paths when needed, sets environment variables, and forwards all remaining arguments to the executable.

### 3.5 Embed the Server in Your Own Agent

When launching the MCP server from another application:

1. Ensure `ExcelMcp.Server[.exe]` is alongside your agent binaries.
2. Spawn it with a workbook argument and capture stdio pipes.
3. Exchange MCP JSON-RPC messages over those pipes.

Example pseudo-code (C#):

```csharp
var start = new ProcessStartInfo
{
    FileName = "ExcelMcp.Server.exe",
    ArgumentList = { "--workbook", workbookPath },
    RedirectStandardInput = true,
    RedirectStandardOutput = true,
    RedirectStandardError = true,
    UseShellExecute = false
};
var process = Process.Start(start);
// Wire the streams into your MCP transport handler
```

## Troubleshooting

### Sidebar Not Visible

**Problem:** The workbook upload sidebar doesn't appear in the web UI.

**Solutions:**
- Hard refresh the browser (Ctrl+Shift+R or Cmd+Shift+R)
- Try in an incognito/private window
- Clear browser cache
- Widen browser window (responsive design hides sidebar on narrow screens < 1024px)

### Workbook Won't Load

**Problem:** Error message when uploading workbook.

**Solutions:**
- Ensure file is a valid `.xlsx` or `.xls` file
- Check that the file isn't open in Excel (file locking)
- Verify file size is under 50MB
- Check console/terminal for detailed error messages

### LLM Not Responding

**Problem:** Queries timeout or return errors.

**Solutions:**
- Verify LM Studio (or your LLM server) is running
- Check the endpoint URL matches (`http://localhost:1234/v1`)
- Ensure a model is loaded in LM Studio
- Test the endpoint: `curl http://localhost:1234/v1/models`
- Check `appsettings.json` has correct `BaseUrl` and `Model`

### MCP Server Crashes

**Problem:** Server process terminates unexpectedly.

**Solutions:**
- Check workbook path is correct and file exists
- Ensure workbook isn't corrupted
- Look for error messages in the terminal
- Try a different workbook to isolate the issue
- Check disk space and permissions

### Table Data Not Rendering

**Problem:** Data appears as text instead of formatted tables.

**Solutions:**
- This should be fixed in the latest version
- Ensure you're running the updated code with CSV-to-HTML conversion
- Check browser console for JavaScript errors

### Permission Denied Errors

**Problem:** Cannot read/write temp files.

**Solutions:**
- Check temp directory permissions (`$env:TEMP` on Windows, `/tmp` on Linux)
- Run with appropriate user permissions
- Specify a custom temp directory with write access

## Advanced Configuration

### Custom Server Path

If auto-detection fails, set the server path explicitly:

```powershell
# Environment variable
$env:EXCEL_MCP_SERVER = "C:\Path\To\ExcelMcp.Server.exe"

# Or in appsettings.json
{
  "ExcelMcp": {
    "ServerPath": "C:/Path/To/ExcelMcp.Server.exe"
  }
}
```

### Custom LLM Configuration

For non-LM Studio endpoints:

```json
{
  "LlmStudio": {
    "BaseUrl": "http://your-server:port/v1",
    "Model": "your-model-name",
    "Temperature": 0.7,
    "MaxTokens": 2048
  }
}
```

### Logging Configuration

Adjust logging levels in `appsettings.json`:

```json
{
  "Serilog": {
    "MinimumLevel": {
      "Default": "Information",
      "Override": {
        "Microsoft": "Warning",
        "System": "Warning",
        "ExcelMcp": "Debug"
      }
    },
    "WriteTo": [
      {
        "Name": "File",
        "Args": {
          "path": "logs/agent-.log",
          "rollingInterval": "Day",
          "outputTemplate": "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"
        }
      }
    ]
  }
}
```

Logs are written to `logs/agent-{date}.log` by default.

## Tips and Best Practices

### Workbook Preparation

- **Name your tables:** Use Excel's Table feature (Ctrl+T) and give tables meaningful names
- **Add headers:** Ensure first row contains column headers
- **Consistent data types:** Keep columns homogeneous (all dates, all numbers, etc.)
- **Remove empty rows:** Clean up blank rows between data sections

### Query Optimization

- **Be specific:** "Show first 10 rows of Sales" vs "show me some data"
- **Use table/sheet names:** Reference exact names from the workbook
- **Start with metadata:** Ask "what sheets exist?" before querying data
- **Limit row counts:** Request specific row limits to avoid overwhelming responses

### Performance

- **Workbook size:** Smaller workbooks (<10MB) load faster
- **Row limits:** Preview 10-50 rows at a time rather than entire tables
- **Search scope:** Use specific search terms to reduce result sets
- **Model selection:** Smaller models (phi-4) are faster than large ones

## Integration Examples

### VS Code MCP Configuration

Add to your `.vscode/mcp.json`:

```json
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

### Python Agent Integration

```python
import subprocess
import json

class ExcelMCPClient:
    def __init__(self, server_path, workbook_path):
        self.process = subprocess.Popen(
            [server_path, "--workbook", workbook_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
    
    def call_tool(self, tool_name, arguments=None):
        request = {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "tools/call",
            "params": {
                "name": tool_name,
                "arguments": arguments or {}
            }
        }
        self.process.stdin.write(json.dumps(request) + "\n")
        self.process.stdin.flush()
        
        response = self.process.stdout.readline()
        return json.loads(response)
    
    def list_structure(self):
        result = self.call_tool("excel-list-structure")
        metadata = json.loads(result["result"]["content"][0]["text"])
        return metadata
    
    def search(self, query):
        return self.call_tool("excel-search", {"query": query})
    
    def preview_table(self, worksheet, rows=10):
        return self.call_tool("excel-preview-table", {
            "worksheet": worksheet,
            "rows": rows
        })

# Usage
client = ExcelMCPClient("./ExcelMcp.Server.exe", "data.xlsx")
structure = client.list_structure()
print(f"Workbook has {len(structure['worksheets'])} sheets")
```

---

## Using the Web Chat Interface (Work in Progress)

**Status:** Functional but needs validation testing. Use CLI agent for critical work.

### Prerequisites
- .NET 9.0 SDK
- Local LLM server (LM Studio, Ollama) on port 1234
- Modern web browser
- For Linux: `libgdiplus` package (`sudo apt install libgdiplus`)

### Quick Start

**Linux/Raspberry Pi:**
```bash
./run-chatweb.sh
# Opens on http://localhost:5001
```

**Windows:**
```pwsh
dotnet run --project src/ExcelMcp.ChatWeb
# Opens on http://localhost:5000
```

### Using the Interface

1. **Load Workbook** - Click "📁 Choose Excel file..." and select .xlsx file
2. **Ask Questions** - Type in chat input (e.g., "Show me the first 10 rows")
3. **View Tables** - Data renders as formatted HTML tables
4. **Switch Workbooks** - Load different files without restarting
5. **Clear History** - Reset conversation, keep workbook

### Features
- ✅ Workbook-agnostic prompts (same as CLI)
- ✅ HTML table rendering
- ✅ Linux/Raspberry Pi support (ARM64)
- ✅ Conversation history
- ✅ Suggested queries

### Troubleshooting
- **Port conflict:** Use `ASPNETCORE_URLS="http://localhost:5001"`
- **Linux:** Install `libgdiplus` if Excel loading fails
- **More help:** See [WEB-CHAT-ROADMAP.md](../WEB-CHAT-ROADMAP.md)

---

## Known Limitations

- **Read-only:** Current version only reads workbooks, no write operations
- **File size:** Very large workbooks (>100MB) may be slow to load
- **Formulas:** Returns calculated values, not formula definitions
- **Charts:** Chart definitions are not exposed through MCP tools
- **Pivot tables:** Pivot table data is not yet supported
- **Protected workbooks:** Password-protected workbooks are not supported

See [docs/FutureFeatures.md](FutureFeatures.md) for planned enhancements.

## 4. Refreshing Bundles After Changes

Whenever you update the source:

1. Run `dotnet test` (recommended).
2. Re-run the relevant packaging scripts for each RID you ship.
3. Distribute the newly generated zip/folder.

```powershell
pwsh -File scripts/package-client.ps1 -Runtime win-x64
pwsh -File scripts/package-server.ps1 -SkipZip
```

Add `-SkipZip` when you only need the staged folder, or supply multiple RIDs in a single invocation.

---
Questions or feedback? Open an issue or reach out via the project’s discussion channels.
