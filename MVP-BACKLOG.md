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
| 2.1 | Add write contracts to `ExcelMcp.Contracts` (`WriteCellRequest`, `WriteRangeRequest`, `WriteResult`) | ⬜ | |
| 2.2 | Implement `ExcelWriteService.cs` in Server/Excel/ using ClosedXML | ⬜ | Includes file backup logic |
| 2.3 | Add MCP tool: `excel-write-cell` | ⬜ | Depends on 1.3 |
| 2.4 | Add MCP tool: `excel-write-range` | ⬜ | Depends on 1.3 |
| 2.5 | Add MCP tool: `excel-create-worksheet` | ⬜ | Low effort add-on |
| 2.6 | Add write functions to `ExcelPlugin.cs` (SkAgent) | ⬜ | |
| 2.7 | Add write plugins to ChatWeb `Services/Plugins/` | ⬜ | |
| 2.8 | Add unit tests for `ExcelWriteService` | ⬜ | |
| 2.9 | Add contract tests for write DTOs | ⬜ | |

---

## Phase 3 — Ollama as Default Provider + Provider Flexibility

> **Why:** Ollama is now the dominant local LLM runtime. LM Studio support stays but Ollama becomes the default. SK version bump to pick up streaming improvements.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 3.1 | Update `appsettings.json` defaults to Ollama (`http://localhost:11434/v1`, model: `llama3.2`) | ⬜ | |
| 3.2 | Add `LlmProviderOptions.cs` with `ProviderType` enum (`Ollama`, `LmStudio`, `OpenAI`) | ⬜ | |
| 3.3 | Update ChatWeb `Program.cs` — auto-detect provider at startup (try Ollama, fall back to LM Studio) | ⬜ | |
| 3.4 | Add provider status indicator to Chat UI | ⬜ | Small UI badge |
| 3.5 | Update `AgentConfiguration.cs` in SkAgent with same Ollama defaults | ⬜ | |
| 3.6 | Bump `Microsoft.SemanticKernel` to latest in both ChatWeb and SkAgent projects | ⬜ | Check for breaking changes |

---

## Phase 4 — MCP Client Integration Configs

> **Why:** Zero-friction onboarding for Claude Desktop, GitHub Copilot, Cursor users.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 4.1 | Create `mcp-config/claude_desktop_config.json` example | ⬜ | |
| 4.2 | Create `.vscode/mcp.json` for GitHub Copilot agent mode | ⬜ | |
| 4.3 | Create `mcp-config/cursor_mcp_config.json` example | ⬜ | |
| 4.4 | Add `scripts/install.sh` — build + copy to Claude Desktop config path (Linux/macOS) | ⬜ | |
| 4.5 | Add `scripts/install.ps1` — build + copy to Claude Desktop config path (Windows) | ⬜ | |
| 4.6 | Update `README.md` QuickStart with per-client setup instructions | ⬜ | |

---

## Phase 5 — Streaming in Blazor UI

> **Why:** Real-time token streaming is the expected UX in 2026. Users shouldn't wait for a full response.

| # | Task | Status | Notes |
|---|------|--------|-------|
| 5.1 | Update `ChatService.cs` to use SK `GetStreamingChatMessageContentsAsync` | ⬜ | |
| 5.2 | Update `Chat.razor` to accumulate tokens and call `StateHasChanged()` per chunk | ⬜ | |
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
