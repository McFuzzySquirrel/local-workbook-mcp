---
ejs:
  type: journey-adr
  version: 1.1
  adr_id: "0001"
  title: MVP Architecture — MCP SDK Migration, Write Operations, and Streaming UI
  date: 2026-03-11
  status: accepted
  session_id: ejs-session-2026-03-11-01
  session_journey: ejs-docs/journey/2026/ejs-session-2026-03-11-01.md

actors:
  humans:
    - id: McFuzzySquirrel
      role: project owner / decision approver
  agents:
    - id: copilot
      role: primary implementation agent
    - id: Explore
      role: codebase assessment sub-agent

context:
  repo: local-workbook-mcp
  branch: feature/mvp-mcp-sdk-migration
---

# Session Journey

Link to the originating session artifact:
- Session Journey: `ejs-docs/journey/2026/ejs-session-2026-03-11-01.md`

# Context

The project had a feature-complete read-only Excel MCP server and Blazor web chat UI, but three significant gaps blocked "true MVP" status:

1. **MCP server was hand-rolled** (custom JSON-RPC over stdio, ~400 LOC). Claude Desktop, GitHub Copilot Agent Mode, Cursor, and VS Code all expect the official `ModelContextProtocol` wire format with spec-compliant tool discovery — the hand-rolled server was not compatible.
2. **No write-back capability** — users could only read Excel data, not modify it.
3. **LM Studio as the only default** — Ollama had overtaken LM Studio as the dominant local LLM runtime. Provider detection was single-target and had no fallback.
4. **No streaming in the Blazor UI** — responses required a full round-trip before anything appeared.

Additional environmental constraints discovered during the session:
- Ubuntu 24.04 (noble) apt only provides .NET 8 and .NET 10 runtimes — not .NET 9. All projects were targeting `net9.0`, making test execution impossible without the runtime.
- `Microsoft.SemanticKernel.Core` 1.26.0/1.28.0 had a critical CVE (GHSA-2ww3-72rp-wpp4) disclosed by `NU1904` during build.

---

# Session Intent

Transform the project into a universal, production-ready MVP that:
- Works as a **pluggable MCP server** for Claude Desktop, GitHub Copilot, Cursor, and VS Code Agent Mode
- Works as a **standalone Blazor web application** with its own LLM integration
- Supports **write operations** (cell update, range write, worksheet creation)
- Defaults to **Ollama** while retaining LM Studio as a fallback
- Provides real-time **streaming responses** in the Blazor UI

# Collaboration Summary

The session proceeded sequentially through 5 planned phases:

1. **Phase 1** (prior session context): Replaced hand-rolled JSON-RPC with `ModelContextProtocol` v1.1.0 SDK. Rewritten `Program.cs` + new `ExcelTools.cs`. Dual-mode workbook path resolution (startup `--workbook` flag OR per-call `workbook_path` parameter). Deleted 4 obsolete files.

2. **Phase 2**: Created `WriteContracts.cs`, `ExcelWriteService.cs` (with auto-timestamped backup), 3 new MCP tools (`excel-write-cell`, `excel-write-range`, `excel-create-worksheet`), SK write plugins for both SkAgent CLI and ChatWeb Blazor.

3. **Phase 3**: Replaced single-target `DetectRunningModelAsync` with cascading `DetectProviderAsync` that tries Ollama (`:11434`) first, then LM Studio (`:1234`). Updated both `appsettings.json` files. Updated SkAgent fallback defaults. Dropped the proposed `ProviderType` enum in favour of transparent auto-detection.

4. **Phase 4**: Created client integration configs for Claude Desktop, Cursor, and VS Code/Copilot Agent Mode. Created cross-platform `install.sh` + `install.ps1`.

5. **Phase 5**: Implemented `StreamQueryAsync` in `ExcelAgentService` using `GetStreamingChatMessageContentsAsync`. Updated `Chat.razor` to iterate chunks and re-render per chunk.

**Pivots:**
- Proposed `ProviderType` enum approach → dropped in favour of cascade auto-detect (simpler, zero-config for users).
- Discovered net9.0 TFM incompatibility at test-run time → bumped all TFMs to `net10.0`.
- Discovered critical CVE during build → immediately upgraded SK to 1.73.0 before committing.

---

# Decision Trigger / Significance

This session warrants an ADR because it contains **multiple cross-cutting architectural decisions** that affect the server's wire protocol, the data mutation model, the LLM provider strategy, and the UI rendering model. These decisions set the foundation for all future development and external integrations.

---

# Considered Options

## Option A — Official MCP SDK (chosen)
Replace hand-rolled JSON-RPC with `ModelContextProtocol` NuGet SDK (`[McpServerToolType]` / `[McpServerTool]` attributes, `AddMcpServer().WithStdioServerTransport().WithTools<T>()`).

## Option B — Keep hand-rolled MCP, fix compatibility manually
Manually align the custom JSON-RPC implementation with the MCP spec. Add proper `initialize` / `tools/list` / `tools/call` response shapes.

---

## Option A — Cascade auto-detect for LLM provider (chosen)
`DetectProviderAsync`: try `[:11434/v1/models, :1234/v1/models]` in order; on first success, override both `SemanticKernelOptions.Model` and `SemanticKernelOptions.BaseUrl`.

## Option B — Explicit `ProviderType` enum + selection UI
Add `ProviderType { Ollama, LmStudio, OpenAI }` to options; let user select in config or at startup.

---

## Option A — In-place mutation of placeholder `ConversationTurn` for streaming (chosen)
`StreamQueryAsync` adds a mutable placeholder turn to `ConversationHistory`, updates `Content` per chunk, finalises with `FormatResponseContent`, syncs to `ContextWindow` manually.

## Option B — Separate streaming state in the Blazor component
Maintain a `streamingContent` variable in `Chat.razor`, show it as a temporary bubble, replace with final turn at end of stream.

---

# Decision

1. **Adopt the official `ModelContextProtocol` NuGet SDK** for all MCP server functionality. All custom JSON-RPC infrastructure is deleted.
2. **Write operations are in-scope for MVP** — `excel-write-cell`, `excel-write-range`, `excel-create-worksheet` with file backup.
3. **Ollama is the new default LLM provider**; LM Studio is retained as a transparent fallback via cascade auto-detection. No user configuration required.
4. **Per-call `workbook_path` parameter on every MCP tool** (dual-mode: startup flag or per-call). This enables both standalone and external-client usage from a single binary.
5. **Streaming via in-place `ConversationTurn` mutation** — `ExcelAgentService.StreamQueryAsync` owns the turn lifecycle; `Chat.razor` only iterates and re-renders.
6. **All TFMs set to `net10.0`** — the .NET version available on this machine's Ubuntu 24.04 apt. No .NET 9 runtime available.
7. **`Microsoft.SemanticKernel` upgraded to 1.73.0** — fixes critical CVE GHSA-2ww3-72rp-wpp4.

---

# Rationale

### MCP SDK over hand-rolled
The official SDK handles all protocol framing, schema generation, tool discovery, and `initialize` handshake automatically. Maintaining compatibility manually would require tracking every SDK release and spec change. The SDK dropped ~400 LOC to ~150 LOC and guarantees compatibility with all current and future MCP clients.

### Write operations in MVP scope
A read-only agent provides limited practical value. Users immediately ask "can I update that cell?" — committing to write-back in Phase 1 avoids a painful later refactor and makes the tool genuinely useful.

### Cascade auto-detect over explicit provider selection
Zero configuration is better UX than accurate configuration. Anyone with Ollama or LM Studio running gets the right provider automatically. The enum approach adds cognitive overhead without user benefit — the underlying SK `AddOpenAIChatCompletion` is provider-agnostic anyway.

### In-place mutation for streaming
Adding a placeholder turn to `ConversationHistory` and mutating it avoids any risk of duplicate entries (AddAssistantTurn is not called at finalization). Blazor sees the same object reference the whole time — no flash, no re-order. The alternative (separate state + replace) creates a transient visual inconsistency. This pattern also keeps all session history logic inside the service, not the component.

### net10.0 TFM
Pragmatic choice — the machine has .NET 10. No benefit to staying on a TFM whose runtime isn't present. SK 1.73.0 supports net8.0+ so no library compatibility issue.

### SK 1.73.0 upgrade
Critical CVE. Non-negotiable. Upgrade triggered immediately upon discovering `NU1904` during build. No breaking API changes were found across the codebase.

---

# Consequences

### Positive
- MCP server is now compatible with Claude Desktop, GitHub Copilot Agent Mode, Cursor, and VS Code Agent Mode out of the box.
- Blazor web app can now read AND write Excel workbooks; files are always backed up before mutation.
- Ollama users get a first-class experience with zero configuration.
- Streaming responses make the Blazor UI feel significantly more responsive — especially important for slower local models.
- All 68 tests pass on net10.0 with no regressions.
- Critical security vulnerability eliminated before any public release.

### Negative / Trade-offs
- Tool class `ExcelTools` constructs `ExcelWorkbookService` and `ExcelWriteService` per call (stateless design required by SDK generic constraint). Minimal overhead for file-based I/O operations.
- Auto-detect adds ~3 seconds to startup in the worst case (two 3-second HTTP timeouts if neither provider is running). Acceptable — only happens once at boot.
- Streaming with `AutoInvokeKernelFunctions` means tool calls happen inside the stream session — the UI shows the spinner for tool resolution, then begins streaming the final response. Not a full token-by-token experience during tool calls. Acceptable for MVP.
- net10.0 is a preview channel on the project's current package set — verified no breaking changes but warrants monitoring on next SK update cycle.

---

# Key Learnings

- `.NET 9` is absent from Ubuntu 24.04 apt. Choose TFMs that match an available runtime on all target machines, or document runtime installation explicitly.
- `dotnet build` surfaces `NU1904`/`NU1903` critical vulnerability warnings — treat them as blocking, not advisory.
- `WithTools<T>()` in the MCP SDK requires a non-static class. `sealed` is the correct modifier.
- `GetStreamingChatMessageContentsAsync` with `AutoInvokeKernelFunctions` correctly handles multi-turn tool-call loops and then streams the final LLM response — no manual tool dispatch needed.
- Blazor streaming: mutate shared state from the async loop + `await InvokeAsync(StateHasChanged)`. Never call `StateHasChanged()` directly from a non-Blazor-dispatcher context.
- MCP stdio transport: ALL logging must go to stderr before the host starts. One log line on stdout breaks the protocol framing.

---

# Agent Guidance

Preferred patterns:
- Use `sealed class` (not `static`) for MCP tool type classes when using `WithTools<T>()`.
- Provider auto-detection cascades are preferred over enum-based explicit selection for local LLM routing.
- For Blazor streaming: create a mutable placeholder turn in session state, mutate in-place per chunk, finalise at end. Do not maintain separate streaming state in the component.
- Check `NU1904`/`NU1903` warnings during every `dotnet build` — upgrade immediately.
- Verify available .NET runtimes before assuming a TFM will work in tests: `dotnet --list-runtimes`.

Anti-patterns:
- Do NOT keep hand-rolled MCP protocol code alongside the SDK — conflicting transport handling.
- Do NOT call `StateHasChanged()` directly inside an `async foreach` streaming loop — use `await InvokeAsync(StateHasChanged)`.
- Do NOT assume `ModelContextProtocol` SDK tool classes can be `static`.
- Do NOT leave `NU1904` vulnerability warnings unaddressed.

---

# Reuse Signals (Optional)

```yaml
reuse:
  patterns:
    - cascade-provider-autodetect
    - mcp-sdk-tool-host-sealed-class
    - blazor-streaming-inplace-mutation
    - mcp-dual-mode-workbook-path
  anti_patterns:
    - static-mcp-tool-class
    - direct-statehschanged-in-async-stream
    - hand-rolled-jsonrpc-over-mcp
  future_considerations:
    - streaming cancel button in Chat UI
    - provider status indicator badge in sidebar
    - write operation unit tests
    - README QuickStart per-client setup section
```
