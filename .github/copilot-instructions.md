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


# Copilot repository instructions (Engineering Journey System)

Follow the Engineering Journey System (EJS) contracts in this repo.

Full agent profile: `.github/agents/ejs-journey.agent.md`

---

## EJS Silent Recording Contract (Always-On)

> **This block is designed to be appended to any repo's existing copilot-instructions.md.**
> It injects recording behavior into whatever agent is currently active — no agent selection needed.
> It does not compete with or override existing agent instructions.

### What You Must Do (Silently, As You Work)

As the active agent in this session, you are responsible for recording your collaboration to the Session Journey file. Do this **automatically as part of your normal work** — do not prompt the user for permission to record.

#### 1. Session Journey File
- If a Session Journey file exists for today under `ejs-docs/journey/YYYY/ejs-session-YYYY-MM-DD-<seq>.md`, update it.
- If none exists, create one from the template at `ejs-docs/journey/_templates/journey-template.md`.
- Use session ID format: `ejs-session-YYYY-MM-DD-<seq>` (2-digit daily sequence).

#### 2. Record Interactions
As each interaction occurs, append to the **Interaction Summary** section:
```
- Human: <prompt / request>
  - Agent [your-name]: <response summary>
  - Outcome: <what changed / what was decided>
```
Attribute every entry by agent name. If you delegate to a sub-agent, record that too.

#### 3. Record Decisions Automatically
When you make or recommend a decision (approach choice, dependency selection, architecture shift, trade-off resolution):
- Add it to the **Decisions Made** section with: decision, reason, alternatives considered, impact
- Do this at the moment the decision occurs — do not wait until session end

#### 4. Record Experiments and Pivots
- When you try something and observe a result, add it to **Experiments / Evidence**
- When the approach changes, add it to **Iteration Log** with what triggered the pivot

#### 5. Record Sub-Agent Work
When you delegate to a sub-agent:
- Log the delegation in the Interaction Summary (what task, which sub-agent, what context)
- After the sub-agent completes, capture its contribution in the **Sub-Agent Contributions** section:
  - Task delegated
  - Decisions made (with rationale)
  - Alternatives considered
  - Outcome
  - Handoff to other agents
- Instruct sub-agents to use the EJS ADR template (`ejs-docs/adr/0000-adr-template.md`) and place ADRs under `ejs-docs/adr/` — do not let sub-agents create conflicting templates or instruction files

#### 6. Record Agent Influence
Track in the **Agent Influence** section:
- Which of your suggestions the human adopted
- Which the human rejected (and why, if stated)
- Human overrides or corrections

### What You Must NOT Do
- Do not prompt the user to "capture this decision" — capture it yourself silently
- Do not interrupt the implementation workflow to ask about recording
- Do not skip recording because the decision seems "obvious"
- Do not fabricate rationale — if the reason is unclear, record "rationale not stated"

### At Session End (wrap up / commit / push / ship)
When the user signals session end:
- Complete all sections with coherent summaries
- Populate all machine extracts (`INTERACTION_EXTRACT`, `DECISIONS_EXTRACT`, `LEARNING_EXTRACT`, `AGENT_GUIDANCE_EXTRACT`, `SUB_AGENT_EXTRACT`)
- Set `decision_detected` field appropriately
- Create an ADR at `ejs-docs/adr/NNNN-<kebab-title>.md` only when the decision rubric triggers (see `.github/agents/ejs-journey.agent.md` for the full rubric)

### Key Principle
Capture context **incrementally throughout the session**, not reconstructed at the end. This produces better documentation by preserving details when they're fresh.

### EJS Database (Required — DB-First Lookup)

Always query the EJS database **before** reading raw markdown files. The database is the primary lookup method; markdown files are a fallback for additional detail.

1. **At session start**, run `python scripts/adr-db.py sync` to refresh the SQLite index.
2. **When referencing past decisions or context**, use database commands first:
   - `python scripts/adr-db.py search <query>` — full-text search across ADRs and journeys
   - `python scripts/adr-db.py summary` — compact overview of all ADRs
   - `python scripts/adr-db.py summary-journeys` — compact overview of all journeys
   - `python scripts/adr-db.py get <adr_id>` — full details for a specific ADR
   - `python scripts/adr-db.py get-journey <session_id>` — full details for a specific journey
3. **Only read the raw markdown files** (`ejs-docs/adr/*.md`, `ejs-docs/journey/**/*.md`) when the database results are insufficient and you need more detail as a fallback.

> **Why DB-first?** Database queries consume far less context than reading full markdown files. The DB returns only relevant snippets and metadata, preserving context budget for actual work.

### Context-Threshold Checkpointing

Do not rely solely on session-end signals to save EJS documentation. Perform **checkpoint saves** proactively during the session to prevent documentation loss if context runs out.

#### When to Checkpoint
Perform an EJS checkpoint save when **any** of the following are true:
- You have accumulated **3 or more unsaved interactions** in your working memory (an interaction is one human prompt and the corresponding agent response)
- You have made a **significant decision** that is not yet written to the journey file
- You are about to start a **large, context-intensive operation** (e.g., reading many files, running complex builds, delegating to sub-agents)
- **5 or more exchanges** have occurred since the last journey file save
- The user has not explicitly ended the session, but substantial work has been completed

#### How to Checkpoint
1. Write all pending interactions, decisions, experiments, and learnings to the Session Journey file
2. Keep the checkpoint lightweight — append to existing sections, do not rewrite the entire file
3. Do **not** populate machine extracts or evaluate the ADR rubric during checkpoints (those are finalization-only)
4. Continue working normally after the checkpoint

#### Principle
Treat each checkpoint as insurance against context loss. If the session were to end unexpectedly after a checkpoint, the journey file should contain a useful (if incomplete) record of work done so far.

Do not claim commands/tests ran unless you observed the output.
