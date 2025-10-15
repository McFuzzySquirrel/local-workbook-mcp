# Excel Local MCP

Local Model Context Protocol (MCP) server and sample client for working with on-disk Excel workbooks. The server exposes worksheet metadata, table previews, and search capabilities via MCP tools and resources so that AI agents can reason over spreadsheet data without uploading it to the cloud.

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
| `excel.list_structure`  | Summarize worksheets, tables, and column headers.               |
| `excel.search`          | Search for rows containing a text query with optional filters.  |
| `excel.preview_table`   | Return a CSV preview of a worksheet or table section.           |

### Resources

The server exposes the workbook plus derived worksheet/table resources using the `excel://` scheme. `resources/read` responses include CSV or JSON previews that agents can consume directly.

## Sample Client

The client launches the server process, performs MCP requests, and prints responses to the console.

```pwsh
# Provide workbook (required) and optional server path
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" list

# Search worksheet or table content
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" search --query "Contoso" --worksheet "Sales"

# Preview a table
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" preview --worksheet "Sales" --table "FY24_Summary" --rows 15

# List exposed resources
src/ExcelMcp.Client/bin/Debug/net9.0/ExcelMcp.Client.exe --workbook "D:/Data/finance.xlsx" resources
```

Client flags:

- `--server` / `EXCEL_MCP_SERVER` – Optional fully qualified path to the server executable. Defaults to the debug build output if present.
- `--workbook` / `-w` / `EXCEL_MCP_WORKBOOK` – Path to the Excel workbook (required).

### Command Reference

- `list` – Calls `excel.list_structure` and prints the textual summary.
- `search` – Calls `excel.search`; accepts `--query`, `--worksheet`, `--table`, `--limit`, `--case-sensitive`.
- `preview` – Calls `excel.preview_table`; accepts `--worksheet`, `--table`, `--rows`.
- `resources` – Performs `resources/list` and prints resource metadata.

## Integrating with Other Agents

Point your MCP-aware agent at the `ExcelMcp.Server` executable, passing the workbook path via `--workbook` or environment variable. After initialization, the agent can discover tools/resources with standard MCP requests and invoke them to pull spreadsheet context into prompts or downstream tools.

## Notes

- ClosedXML loads workbooks into memory; very large files may incur higher memory usage.
- Search defaults to the first 20 matches; increase via `limit` but beware of large responses.
- Resource previews return CSV or JSON snippets to keep responses small; fetch the underlying workbook directly if you need the full data.
