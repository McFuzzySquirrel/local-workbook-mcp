---
ejs:
  type: journey-adr
  version: 1.1
  adr_id: "0002"
  title: CLI Debug Client — MCP SDK Protocol Compliance and Write Command Parity
  date: 2026-03-11
  status: accepted
  session_id: ejs-session-2026-03-11-01
  session_journey: ejs-docs/journey/2026/ejs-session-2026-03-11-01.md

actors:
  humans:
    - id: McFuzzySquirrel
      role: project owner
  agents:
    - id: copilot
      role: implementation agent

context:
  repo: local-workbook-mcp
  branch: feature/mvp-mcp-sdk-migration
---

# Session Journey

Link to the originating session artifact:
- Session Journey: `ejs-docs/journey/2026/ejs-session-2026-03-11-01.md`

# Context

After completing the full MVP implementation (Phases 1–5), `ExcelMcp.Client` — the low-level CLI debug client — was out of sync with the updated MCP server in three ways:

1. **Missing `notifications/initialized`** — The MCP spec requires the client to send this notification after receiving the server's `initialize` response. The SDK server (`ModelContextProtocol` v1.1.0) enters a ready state only after this notification. Without it, `tools/list` and `tools/call` requests time out or are silently ignored. The original hand-rolled server was more lenient; the SDK server is spec-strict.

2. **Hardcoded `net9.0` dev path** — `FindServerExecutable()` contained `bin/Debug/net9.0/ExcelMcp.Server.exe`. All TFMs were bumped to `net10.0` in Phase 3, so the auto-discovery path was broken on this machine.

3. **No write commands** — Phase 2 added three MCP write tools (`excel-write-cell`, `excel-write-range`, `excel-create-worksheet`). The CLI debug client had no commands for them, creating a feature gap.

`ExcelMcp.Client` is a thin diagnostic/debugging tool (not an end-user product), but it should stay in lock-step with the server's tool surface so developers can test any tool from the command line.

---

# Session Intent

Bring `ExcelMcp.Client` into full compliance with the SDK server and full parity with its tool surface:
- Fix the `notifications/initialized` protocol compliance gap
- Fix the broken dev server auto-discovery path
- Add `write-cell`, `write-range`, `create-worksheet` command handlers

---

# Collaboration Summary

Reviewed both `ExcelMcp.Client` source files (`Program.cs`, `Mcp/McpProcessClient.cs`, `Mcp/JsonRpcClient.cs`) to understand the existing structure before making changes. No refactoring was done — changes were minimal and scoped to the exact gaps.

`InitializeAsync` in `McpProcessClient.cs` was extended to issue `notifications/initialized` via the existing `SendNotificationAsync` path immediately after a successful `initialize` response. `protocolVersion: "2024-11-05"` was also added to the initialize parameters to be explicit about the version being negotiated.

`FindServerExecutable()` in `Program.cs` was updated to probe the `net10.0` debug output path on Linux (no `.exe`) and Windows (`.exe`), replacing the single `net9.0` candidate.

Three new command cases were added to the `switch` block (`write-cell`, `write-range`, `create-worksheet`) with corresponding static handler functions. The `--data` argument for `write-range` accepts a raw JSON array string. `ShowHelp()` was split into **Read commands** / **Write commands** sections.

---

# Decision Trigger / Significance

This warrants an ADR because the `notifications/initialized` gap represents a **protocol compliance decision** with consequences for anyone running the CLI client against the SDK server: without it, the client appears to connect successfully but all subsequent calls fail silently. The decision to keep the client's own hand-rolled JSON-RPC layer (rather than switching it to the official SDK) also needs to be recorded.

---

# Considered Options

## Option A — Keep hand-rolled `McpProcessClient`, fix the protocol gap (chosen)
Send `notifications/initialized` via the existing `JsonRpcClient.SendNotificationAsync` path. Keep the custom JSON-RPC client intact.

## Option B — Replace `McpProcessClient` with the official `ModelContextProtocol` client SDK
Use `McpClientFactory` from `ModelContextProtocol` NuGet in the CLI client. Eliminates the custom JSON-RPC layer entirely.

---

# Decision

1. **Keep `McpProcessClient` hand-rolled** — `ExcelMcp.Client` is a thin diagnostic tool. The official client SDK adds significant dependency weight for a tool whose only purpose is manual testing. The hand-rolled layer is ~150 LOC and well-understood.
2. **Fix `notifications/initialized`** — Send it immediately after each successful `initialize` response. Add `protocolVersion: "2024-11-05"` to the `initialize` request for explicitness.
3. **Fix `FindServerExecutable`** — Probe `net10.0` paths (Linux + Windows) instead of `net9.0`.
4. **Add write commands** — `write-cell`, `write-range`, `create-worksheet` with `--sheet`, `--cell`/`--range`, `--value`/`--data` arguments.

---

# Rationale

### Keep hand-rolled client
`ExcelMcp.Client` is not shipped to end users. It exists so developers can poke at the server from the terminal. Pulling in the full MCP client NuGet adds indirection and package surface without adding value for this use case. The fix to `notifications/initialized` is a 3-line change; it does not justify a full rewrite.

### `notifications/initialized` is mandatory per spec
The MCP spec (2024-11-05) states: "After the client sends the initialize request and receives a successful response, it MUST send a `notifications/initialized` notification before sending any other requests." The SDK server enforces this strictly. Any client omitting it will behave erratically when connected to a spec-compliant server.

### Write command parity
The CLI client's value is proportional to how many tools it can exercise. Adding the three write commands makes it a complete testing harness for all 7 tools the server exposes.

---

# Consequences

### Positive
- `ExcelMcp.Client` can now successfully call any of the 7 MCP tools against the SDK server without hanging.
- Developers can test write operations from the command line without spinning up the Blazor UI.
- Dev path auto-discovery works on this machine after the `net9.0` → `net10.0` TFM bump.

### Negative / Trade-offs
- `McpProcessClient` remains hand-rolled and must be kept manually in sync with future MCP spec changes. The spec-compliance surface is small (initialize handshake + tool call) but requires ongoing vigilance.
- The `--data` argument for `write-range` requires the caller to construct a JSON array string manually. No ergonomic CLI DSL (e.g., `--cell A1 --value x --cell B1 --value y`). Acceptable for a diagnostic tool.

---

# Key Learnings

- The `ModelContextProtocol` SDK server is strict about the `notifications/initialized` handshake. Lenient hand-rolled servers that skipped this worked; spec-compliant servers do not. Always send it.
- When bumping TFMs, scan for hardcoded TFM strings in dev path resolution, install scripts, and documentation — not just `.csproj` files.
- A hand-rolled client and an SDK server can coexist cleanly as long as the client sends the correct protocol handshake messages.

---

# Agent Guidance

Preferred patterns:
- Always send `notifications/initialized` after `initialize` in any hand-rolled MCP client.
- Include `protocolVersion` in the `initialize` request params to make version negotiation explicit.
- When adding new MCP tools server-side, update `ExcelMcp.Client/Program.cs` in the same commit or PR.

Anti-patterns:
- Do NOT hardcode TFM strings like `net9.0` in dev path discovery — derive from build output or use glob probing.
- Do NOT test a hand-rolled client against a lenient server and assume it will work against a spec-compliant one.

---

# Reuse Signals (Optional)

```yaml
reuse:
  patterns:
    - mcp-handshake-notifications-initialized
    - tfm-agnostic-dev-path-discovery
  anti_patterns:
    - hardcoded-tfm-in-path-resolution
    - skip-notifications-initialized
  future_considerations:
    - migrate McpProcessClient to official SDK client when client complexity grows
    - ergonomic multi-cell --cell/--value flag pairs for write-range
```
