# Excel Local MCP

Local Model Context Protocol server, CLI client, and chat web UI for working with on-disk Excel workbooks. The server exposes worksheet metadata, table previews, and search capabilities via MCP tools and resources. Everything runs locally, ships with self-contained binaries, and exposes a consistent MCP surface so agents can inspect spreadsheets without uploading them.

> Need the full walkthrough? See [docs/UserGuide.md](docs/UserGuide.md).

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

- `src/ExcelMcp.Server` – Stdio JSON-RPC MCP server that indexes a workbook.
- `src/ExcelMcp.Client` – Command-line tool that starts the server and calls MCP tools/resources.
- `src/ExcelMcp.ChatWeb` – ASP.NET front end that talks to the server and renders the chat UI.
- `src/ExcelMcp.Contracts` – Shared data contracts.
- `docs/UserGuide.md` – Extended walkthrough covering setup, workflows, and troubleshooting.
- `docs/FutureFeatures.md` – Forward-looking ideas we plan to explore.

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

> **Note:** The chat web app expects [LM Studio](https://lmstudio.ai/) (or any OpenAI-compatible local server) to be running on `http://localhost:1234` with the `phi-4-mini-reasoning` model downloaded and loaded. Update `LlmStudio:BaseUrl` and `LlmStudio:Model` in `appsettings.json` if you host a different endpoint/model. See the user guide for setup steps.

```pwsh
# Run the stdio MCP server directly
dotnet run --project src/ExcelMcp.Server -- --workbook "D:/Data/sample.xlsx"

# Use the CLI client (prompts for missing workbook/server values)
dotnet run --project src/ExcelMcp.Client -- list

# Launch the chat UI (serves wwwroot at http://localhost:5000)
dotnet run --project src/ExcelMcp.ChatWeb
```

## Self-Contained Bundles

Each app has a dedicated packaging script that publishes single-file executables, copies launch helpers, and writes a README into `dist/<rid>/<AppName>`:

```pwsh
pwsh -File scripts/package-server.ps1   # ExcelMcp.Server
pwsh -File scripts/package-client.ps1   # ExcelMcp.Client
pwsh -File scripts/package-chatweb.ps1  # ExcelMcp.ChatWeb (includes wwwroot)
```

Pass `-Runtime` (e.g., `linux-x64`) or `-SkipZip` as needed. The generated folders include `run-*.ps1`, `run-*.sh`, and `run-*.bat` wrappers that prompt for workbook/server paths and start the bundled executable.

## Integrate with Agents

Point your MCP-capable agent at the packaged server or use the CLI/web app launchers to negotiate workbook/server paths. All tools (`excel-list-structure`, `excel-search`, `excel-preview-table`) and resources (`excel://` URIs) follow the MCP spec, so they work with OpenAI agents, MCP bridges, or any compatible orchestrator.

Refer to [docs/UserGuide.md](docs/UserGuide.md) for detailed workflow examples, environment variables, and troubleshooting notes.

## Acknowledgements

This project relies on the excellent ClosedXML library (https://github.com/ClosedXML/ClosedXML) as its core engine for reading and writing Excel workbooks (.xlsx files).
ClosedXML provides a clean, intuitive API built on top of the OpenXML SDK, allowing this MCP server to interact with Excel data without requiring Microsoft Excel or any COM automation.

Every workbook operation exposed through this MCP server, such as reading rows, writing cell values, or adding data, is powered internally by ClosedXML.

A huge thank you to the ClosedXML maintainers and contributors for their ongoing work on one of the most reliable and developer-friendly Excel libraries in the .NET ecosystem.
Your project makes it possible for tools like this to exist and run cross-platform in lightweight environments