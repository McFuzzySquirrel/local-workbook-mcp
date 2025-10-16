# Excel Local MCP

Local Model Context Protocol (MCP) server and sample client for working with on-disk Excel workbooks. The server exposes worksheet metadata, table previews, and search capabilities via MCP tools and resources so that AI agents can reason over spreadsheet data without uploading it to the cloud.

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

- Support filtered range previews (e.g., formulas vs. values, pivot expansions).
- Implement write-back tools to update cells, add worksheets, or annotate findings.
- Expose analytics such as value distributions, outlier detection, or chart generation.
- Add WebSocket/HTTP transports so the server can run behind a relay or container orchestration platform.

## Projects

- `src/ExcelMcp.Server` – MCP server that indexes an Excel workbook with ClosedXML and exposes tools/resources over stdio JSON-RPC.
- `src/ExcelMcp.Client` – Command-line client that launches the server process, performs the MCP handshake, and demonstrates tool and resource usage.
- `src/ExcelMcp.Contracts` – Shared contracts for workbook metadata and search responses.

## Prerequisites

- .NET SDK 9.0 or newer (`dotnet --version` should report 9.x)
- Windows, macOS, or Linux capable of running .NET console applications
- A local `.xlsx` workbook to inspect

## Build

```pwsh
# From the repository root
dotnet build
```

## Running the MCP Server

Supply the target workbook path via command-line or environment variable:

```pwsh
# Option 1: command-line
src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server.exe --workbook "D:/Data/finance.xlsx"

# Option 2: environment variable
$env:EXCEL_MCP_WORKBOOK = "D:/Data/finance.xlsx"
src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server.exe
```

The server communicates over stdio following the MCP JSON-RPC framing rules. Terminate with `Ctrl+C` or send a `shutdown` request followed by process exit.

### Tools

| Tool name               | Purpose                                                         |
|-------------------------|-----------------------------------------------------------------|
| `excel-list-structure`  | Summarize worksheets, tables, and column headers.               |
| `excel-search`          | Search for rows containing a text query with optional filters.  |
| `excel-preview-table`   | Return a CSV preview of a worksheet or table section.           |

### Resources

The server exposes the workbook plus derived worksheet/table resources using the `excel://` scheme. `resources/read` responses include CSV or JSON previews that agents can consume directly.

## Sample Client

The client launches the server process, performs MCP requests, and prints responses to the console.

```pwsh
# Provide workbook (required) and optional server path
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" list

# Search worksheet or table content
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" search --query "My Company" --worksheet "Sales"

# Preview a table
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" preview --worksheet "Sales" --table "FY25_Summary" --rows 15

# List exposed resources
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" resources
```

Client flags:

- `--server` / `EXCEL_MCP_SERVER` – Optional fully qualified path to the server executable. Defaults to the debug build output if present.
- `--workbook` / `-w` / `EXCEL_MCP_WORKBOOK` – Path to the Excel workbook (required).

### Command Reference

- `list` – Calls `excel-list-structure` and prints the textual summary.
- `search` – Calls `excel-search`; accepts `--query`, `--worksheet`, `--table`, `--limit`, `--case-sensitive`.
- `preview` – Calls `excel-preview-table`; accepts `--worksheet`, `--table`, `--rows`.
- `resources` – Performs `resources/list` and prints resource metadata.

## Integrating with Other Agents

Point your MCP-aware agent at the `ExcelMcp.Server` executable, passing the workbook path via `--workbook` or environment variable. After initialization, the agent can discover tools/resources with standard MCP requests and invoke them to pull spreadsheet context into prompts or downstream tools.

### OpenAI Agent Example

The `examples/openai_agent.py` script shows how to register the Excel MCP server as a tool on an OpenAI agent using the Python SDK.

1. Publish or build the server:
	```pwsh
	dotnet publish src/ExcelMcp.Server/ExcelMcp.Server.csproj -c Release
	```
2. Set environment variables so the script can find the executable and workbook:
	```pwsh
	$env:EXCEL_MCP_SERVER_PATH = "D:/GitHub Projects/local-workbook-mcp/src/ExcelMcp.Server/bin/Release/net9.0/ExcelMcp.Server.exe"
	$env:EXCEL_MCP_WORKBOOK = "D:/Downloads/sampledata.xlsx"
	$env:OPENAI_API_KEY = "sk-..."
	```
3. Install the OpenAI Python package (version 1.42 or newer):
	```pwsh
	pip install --upgrade openai
	```
4. Run the sample and observe tool invocations in the console:
	```pwsh
	python examples/openai_agent.py
	```

The script creates an agent, registers the MCP transport with `excel-list-structure`, `excel-search`, and `excel-preview-table`, and asks the agent to summarize workbook structure. Modify the prompt or follow-up messages to issue searches or table previews.

## Notes

- ClosedXML loads workbooks into memory; very large files may incur higher memory usage.
- Search defaults to the first 20 matches; increase via `limit` but beware of large responses.
- Resource previews return CSV or JSON snippets to keep responses small; fetch the underlying workbook directly if you need the full data.

## Self-Contained Distribution

Package self-contained binaries so end users do not need the .NET SDK:

```pwsh
pwsh -File scripts/package.ps1            # Produces dist/win-x64 and excel-mcp-win-x64.zip
pwsh -File scripts/package.ps1 -SkipZip   # Skip archive creation
pwsh -File scripts/package.ps1 -Runtime win-arm64
```

The script publishes both the server and client as single-file executables, adds helper launch scripts (`run-client.ps1`, `run-client.bat`, `run-server.ps1`), and drops a quick-start README into `dist/<runtime>`. Send the generated folder or zip to end users; they can double-click `run-client.ps1`, provide a workbook path, and start using the tools immediately. When they need to switch workbooks, re-run the launcher with a different `-WorkbookPath` value.

### Using the bundle on another machine

1. Copy `excel-mcp-<runtime>.zip` (or the entire `dist/<runtime>` folder) to the destination machine.
2. Extract the archive anywhere the user has write permission, for example `C:\Tools\ExcelMcp`.
3. Launch the client with whichever entry point fits the environment:

   ```pwsh
   # PowerShell (asks for the workbook path if omitted)
   ./run-client.ps1

   # Windows Command Prompt shortcut
   run-client.bat "C:\Data\workbook.xlsx"
   ```

4. Follow the prompts; the launcher ensures the bundled server is started with the provided workbook path and keeps the console open for tool commands.

The server launcher (`run-server.ps1`) is included for scenarios where you want to host the MCP server separately and connect with another client. No additional prerequisites are needed beyond Windows PowerShell 5.1+ or PowerShell Core; the .NET runtime is bundled with the executables. If ExecutionPolicy blocks the PowerShell script, start PowerShell with `-ExecutionPolicy Bypass` or `Unblock-File .\run-client.ps1` after extraction.

### MSI Installer (WiX)

To ship a full Windows installer, we include a WiX v4 project that wraps the self-contained bundle into an `.msi`:

```pwsh
dotnet tool restore                      # restores the wix CLI from .config/dotnet-tools.json
pwsh -File scripts/package-wix.ps1       # builds the bundle and creates dist/excel-mcp-win-x64.msi
```

The MSI installs under `Program Files\Excel Local MCP`, adds a Start Menu shortcut pointing to `run-client.bat`, and registers an uninstaller entry. Adjust the `-Runtime` or `-Configuration` parameters to produce installers for other target architectures.
