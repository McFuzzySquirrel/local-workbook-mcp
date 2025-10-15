# Excel Workbook MCP Server

Expose a local Excel workbook as MCP-style resources, tools, query endpoints, and a live diff WebSocket.

Features:
- Parse workbook into sheets and inferred tables using NPOI.
- REST endpoints: /mcp/tools, /mcp/resources, /mcp/query, /mcp/diff, /workbook.
- WebSocket /ws broadcasting JSON ExcelChangeEvent diffs on file changes (debounced with exponential backoff).
- Manifest mcp-manifest.json describing capabilities.
 - Write tools: setCell, appendRow via POST /mcp/tools/invoke.

Run:
- Requires .NET 9 SDK installed.
- Configure ExcelFile in appsettings.json or via environment variable.

Endpoints:
- GET /mcp/tools
- GET /mcp/resources
- GET /mcp/query?sheet=Sheet1
- GET /mcp/query?sheet=Sheet1&table=Table_1
- GET /mcp/diff
- GET /workbook
- WS /ws

Write tools (POST /mcp/tools/invoke):
- setCell: { name: "setCell", args: { sheetName: "Sheet1", address: "B2", value: "Hello" } }
- appendRow: { name: "appendRow", args: { sheetName: "Sheet1", values: "A,B,C" } }

## MCP stdio adapter

For VS Code MCP clients that speak JSON-RPC over stdio, use the adapter console app:

- Project: `ExcelMcp.Adapter`
- Excel file is taken from `ExcelFile` env var or the first CLI argument.
- Methods exposed:
	- tools/listSheets, tools/listTables, tools/getSheetData, tools/getTableData
	- tools/setCell, tools/appendRow

Run example (PowerShell):

```powershell
$env:ExcelFile = "$env:TEMP\mcp-demo.xlsx"
dotnet run --project .\ExcelMcp.Adapter\ExcelMcp.Adapter.csproj
```

### Sample MCP client configs

- VS Code MCP tool config: `examples/vscode-mcp-tool.json`
- OpenAI-style assistant tool config: `examples/openai-assistant-tool.json`

Update the ExcelFile path to point at your workbook. Both use stdio transport and spawn the adapter via `dotnet run`.

#### VS Code .vscode/mcp.json examples

Option A — dotnet run (requires .NET SDK):

```json
{
	"servers": {
		"excel-workbook-mcp": {
			"type": "stdio",
			"command": "dotnet",
			"args": [
				"run",
				"--project",
				"${workspaceFolder}/ExcelMcpServer/ExcelMcp.Adapter/ExcelMcp.Adapter.csproj"
			],
			"env": {
				"ExcelFile": "${input:excel-file}"
			}
		}
	},
	"inputs": [
		{
			"type": "promptString",
			"id": "excel-file",
			"description": "Full path to the Excel workbook",
			"default": "C:/Users/you/Documents/agents/data.xlsx"
		}
	]
}
```

Option B — published binary (no SDK):

```json
{
	"servers": {
		"excel-workbook-mcp-bin": {
			"type": "stdio",
			"command": "${workspaceFolder}/ExcelMcpServer/dist/adapter/ExcelMcp.Adapter.exe",
			"args": [],
			"env": {
				"ExcelFile": "${input:excel-file}"
			}
		}
	},
	"inputs": [
		{
			"type": "promptString",
			"id": "excel-file",
			"description": "Full path to the Excel workbook",
			"default": "C:/Users/you/Documents/agents/data.xlsx"
		}
	]
}
```

## Binary distribution (no dotnet run)

Publish the adapter as a single-file binary and reference it directly in MCP configs:

```powershell
pwsh -File .\ExcelMcpServer\scripts\publish-adapter.ps1 -Runtime win-x64 -Configuration Release -OutDir .\ExcelMcpServer\dist\adapter
```

Then use:
- VS Code config: `examples/vscode-mcp-tool-binary.json` (update command path and ExcelFile)
- OpenAI config: `examples/openai-assistant-tool-binary.json`