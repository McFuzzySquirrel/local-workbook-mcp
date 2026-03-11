# local-workbook-mcp MVP Backlog

> **Branch:** `feature/mvp-mcp-sdk-migration`  
> **Started:** 2026-03-11  
> **Goal:** Transform the project into a true, universal MVP — pluggable with any MCP-compatible client (Claude Desktop, GitHub Copilot, Cursor) AND usable as a standalone Blazor web app with write capability and Ollama support.

---

## Status Legend

| Symbol | Meaning |
|--------|---------|
| ⬜ | Not started |
| 🔄 | In progress |
| ✅ | Complete |
| ❌ | Blocked / won't do |

---

## Phase 1 — Official MCP SDK Migration

> **Why:** The MCP server is hand-rolled (custom JSON-RPC). The official `ModelContextProtocol` NuGet SDK gives full spec compliance and compatibility with Claude Desktop, GitHub Copilot, Cursor, and VS Code Agent Mode.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 1.1 | Add `ModelContextProtocol` + `Microsoft.Extensions.Hosting` NuGet to Server csproj | ✅ | v1.1.0 + v9.0.3 |
| 1.2 | Rewrite `Program.cs` using `AddMcpServer().WithStdioServerTransport()` | ✅ | Logging to stderr; workbook path optional at startup |
| 1.3 | Replace `McpServer.cs` — annotate tool handlers with `[McpServerTool]` | ✅ | New `ExcelTools.cs` with `[McpServerToolType]` + 4 tools |
| 1.4 | Delete `JsonRpcTransport.cs`, `JsonRpcMessage.cs`, `McpModels.cs`, `JsonOptions.cs` (SDK replaces all) | ✅ | First 4 deleted; `JsonOptions.cs` retained (used by ExcelTools) |
| 1.5 | Add `workbook_path` parameter to each tool so external clients can specify files dynamically | ✅ | Dual-mode: per-call param OR `EXCEL_MCP_WORKBOOK` env var |
| 1.6 | Verify Claude Desktop can connect and call `excel-list-structure` | ⬜ | Manual smoke test — needs Claude Desktop installed |
| 1.7 | Verify GitHub Copilot agent mode can connect via `.vscode/mcp.json` | ⬜ | Manual smoke test — needs Phase 4 configs first |
| 1.8 | Update Server unit tests to use SDK in-process test harness | ⬜ | Existing tests still pass; deeper SDK test integration is Phase 1.5+ |

---

## Phase 2 — Write Operations

> **Why:** Users need to make changes, not just read. Minimum viable write: update a cell, write a range. Always backup the file before mutation.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 2.1 | Add write contracts to `ExcelMcp.Contracts` (`WriteCellRequest`, `WriteRangeRequest`, `WriteResult`) | ✅ | `WriteContracts.cs` — 5 records |
| 2.2 | Implement `ExcelWriteService.cs` in Server/Excel/ using ClosedXML | ✅ | Auto-backup w/ timestamp |
| 2.3 | Add MCP tool: `excel-write-cell` | ✅ | In `ExcelTools.cs` |
| 2.4 | Add MCP tool: `excel-write-range` | ✅ | In `ExcelTools.cs` |
| 2.5 | Add MCP tool: `excel-create-worksheet` | ✅ | In `ExcelTools.cs` |
| 2.6 | Add write functions to `ExcelPlugin.cs` (SkAgent) | ✅ | `write_cell`, `write_range`, `create_worksheet` |
| 2.7 | Add write plugins to ChatWeb `Services/Plugins/` | ✅ | `WorkbookWritePlugin.cs` |
| 2.8 | Add unit tests for `ExcelWriteService` | ⬜ | Future |
| 2.9 | Add contract tests for write DTOs | ⬜ | Future |

---

## Phase 3 — Ollama as Default Provider + Provider Flexibility

> **Why:** Ollama is now the dominant local LLM runtime. LM Studio support stays but Ollama becomes the default. SK version bump to pick up streaming improvements.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 3.1 | Update `appsettings.json` defaults to Ollama (`http://localhost:11434/v1`, model: `llama3.2`) | ✅ | Both appsettings.json + appsettings.Development.json |
| 3.2 | Add `LlmProviderOptions.cs` with `ProviderType` enum (`Ollama`, `LmStudio`, `OpenAI`) | ❌ | Replaced by cascade auto-detect; enum not needed |
| 3.3 | Update ChatWeb `Program.cs` — auto-detect provider at startup (try Ollama, fall back to LM Studio) | ✅ | `DetectProviderAsync` tries `:11434` first, then `:1234` |
| 3.4 | Add provider status indicator to Chat UI | ⬜ | Nice-to-have future |
| 3.5 | Update `AgentConfiguration.cs` in SkAgent with same Ollama defaults | ✅ | Fallback defaults + `DetectRunningModel` cascade updated |
| 3.6 | Bump `Microsoft.SemanticKernel` to latest in both ChatWeb and SkAgent projects | ✅ | 1.28.0/1.26.0 → 1.73.0 (fixed GHSA-2ww3-72rp-wpp4 CVE) |

---

## Phase 4 — MCP Client Integration Configs

> **Why:** Zero-friction onboarding for Claude Desktop, GitHub Copilot, Cursor users.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 4.1 | Create `mcp-config/claude_desktop_config.json` example | ✅ | |
| 4.2 | Create `.vscode/mcp.json` for GitHub Copilot agent mode | ✅ | Uses `${workspaceFolder}` var |
| 4.3 | Create `mcp-config/cursor_mcp_config.json` example | ✅ | |
| 4.4 | Add `scripts/install.sh` — build + publish binaries (Linux/macOS) | ✅ | |
| 4.5 | Add `scripts/install.ps1` — build + publish binaries (Windows) | ✅ | |
| 4.6 | Update `README.md` QuickStart with per-client setup instructions | ⬜ | |

---

## Phase 5 — Streaming in Blazor UI

> **Why:** Real-time token streaming is the expected UX in 2026. Users shouldn't wait for a full response.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 5.1 | Update agent service to use SK `GetStreamingChatMessageContentsAsync` | ✅ | `ExcelAgentService.StreamQueryAsync` — mutates placeholder turn in-place |
| 5.2 | Update `Chat.razor` to accumulate tokens and call `StateHasChanged()` per chunk | ✅ | `await InvokeAsync(StateHasChanged)` — first 10 chunks + every 8 thereafter |
| 5.3 | Add streaming cancel button to UI | ⬜ | Nice-to-have |

---

## Out of Scope (Phase 2+)

- Mobile UI polish
- Formula evaluation / what-if scenarios
- Advanced analytics (outlier detection, distributions, charts)
- Slicer / calculated field support in pivot analysis
- Alternative transports (WebSocket/HTTP)

---

## Decisions Log

| Date | Decision | Rationale |
|------|----------|-----------|
| 2026-03-11 | Migrate to official MCP SDK | Unlocks Claude Desktop, Copilot, Cursor, VS Code Agent Mode compatibility |
| 2026-03-11 | Write-back included in MVP | Users expect to make changes; minimum = write-cell + write-range with file backup |
| 2026-03-11 | Ollama as default provider | Dominant local LLM runtime in 2026; LM Studio retained as fallback |
| 2026-03-11 | Each MCP tool accepts `workbook_path` param | External clients (Claude Desktop, Copilot) need to specify files dynamically |
| 2026-03-11 | Blazor UI stays as standalone experience | MCP client configs are additive, not a replacement |
| 2026-03-11 | Cascade auto-detect (Ollama → LM Studio) instead of ProviderType enum | Simpler UX; zero config for anyone with Ollama or LM Studio running |
| 2026-03-11 | Bump all TFMs to net10.0 | Ubuntu 24.04 apt only has .NET 8 and .NET 10 runtimes; net9.0 tests couldn't execute |
| 2026-03-11 | Upgrade SK to 1.73.0 | Fix critical CVE GHSA-2ww3-72rp-wpp4 in SK.Core 1.26.0/1.28.0 |
| 2026-03-11 | Streaming via in-place placeholder turn mutation | No duplicate history entries; Blazor component sees live updates without extra state |
