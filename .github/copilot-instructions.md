# local-workbook-mcp Development Guidelines

## Project Overview
This is a **Local-First, Agentic Architecture** built on the **Model Context Protocol (MCP)** using **.NET 9 (C# 13)**. It enables AI agents to interact with local Excel workbooks via a standardized protocol.

### Core Architecture
- **ExcelMcp.Server**: Standalone process acting as the MCP Server. Uses `ClosedXML` to read Excel files. Exposes tools/resources via JSON-RPC over Stdio.
- **ExcelMcp.ChatWeb**: Blazor Server app acting as the MCP Client. Hosts **Semantic Kernel** to orchestrate AI interactions.
- **ExcelMcp.Client**: CLI tool (Spectre.Console) for direct MCP interaction and debugging.
- **ExcelMcp.Contracts**: Shared library containing `sealed record` DTOs used by both Server and Client.

## Critical Workflows
- **Build**: `dotnet build` (Solution root)
- **Run Web App**: `./run-chatweb.sh` (Starts Server & Web Client)
- **Run CLI**: `dotnet run --project src/ExcelMcp.Client`
- **Generate Test Data**: `./scripts/create-sample-workbook.ps1` (Creates `test-data/ProjectTracking.xlsx`)
- **Package**: Use `scripts/package-*.ps1` for distribution.

## Code Conventions (C# 13 / .NET 9)
- **DTOs**: Use `sealed record` with primary constructors for all data contracts.
  - *Example*: `public sealed record WorkbookMetadata(string Name, long SizeBytes);`
- **Collections**: Prefer `IReadOnlyList<T>` and `IReadOnlyDictionary<K,V>` for public properties.
- **Async**: All I/O and MCP operations must be `async/await`.
- **Logging**: Use `Serilog` with structured logging.
- **Namespaces**: Use file-scoped namespaces (`namespace ExcelMcp.Server;`).

## Integration Patterns
- **MCP Tools**: Implemented in `ExcelMcp.Server`. Map 1:1 to Excel operations (e.g., `read_worksheet`, `pivot_table`).
- **SK Plugins**: Implemented in `ExcelMcp.SkAgent`. Wrap MCP tool calls for the LLM.
- **Communication**: The Web/CLI apps spawn the Server process and communicate via Stdio. Do not try to reference Server code directly from Client apps; use `ExcelMcp.Contracts`.

## Key Files
- `src/ExcelMcp.Contracts/ResourceContracts.cs`: Defines the data shape exchanged between processes.
- `src/ExcelMcp.Server/Program.cs`: Entry point for the MCP Server.
- `src/ExcelMcp.SkAgent/ExcelAgent.cs`: Semantic Kernel orchestration logic.
- `docs/Architecture.md`: Detailed architectural blueprint and diagrams.
