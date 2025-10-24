<!--
SYNC IMPACT REPORT
==================
Version change: TEMPLATE → 1.0.0
Modified principles: N/A (initial constitution)
Added sections:
  - Core Principles (7 principles)
  - Technology Stack & Architecture Constraints
  - Security & Privacy Requirements
  - Governance
Removed sections: None
Templates requiring updates:
  ✅ plan-template.md (Constitution Check section aligns with principles)
  ✅ spec-template.md (requirements align with privacy-first and testability principles)
  ✅ tasks-template.md (test-first and incremental delivery principles reflected)
Follow-up TODOs: 
  - RATIFICATION_DATE marked as TODO (original adoption date unknown)
-->

# Excel Local MCP Constitution

## Core Principles

### I. Privacy-First Local Operation

All data processing MUST occur locally on the user's device. The system MUST NOT transmit spreadsheet content, user queries, or derived insights to external services without explicit user consent. Network connectivity is optional, not required.

**Rationale**: Financial and regulated spreadsheets often cannot leave the device. Local-first architecture ensures compliance with privacy requirements and enables offline operation.

**Enforcement**: All features MUST be testable in airplane mode. Any network call MUST be behind an opt-in flag with explicit user approval in the UI.

### II. MCP Protocol Compliance

All agent-facing functionality MUST be exposed through compliant MCP tools and resources following the Model Context Protocol specification. Tools MUST accept structured JSON inputs and return structured JSON outputs with consistent error formats.

**Rationale**: MCP standardization enables any compliant client (Claude Desktop, VS Code extensions, custom agents) to consume Excel data without custom integration code.

**Enforcement**: Every new capability MUST be implemented as an MCP tool or resource. Direct API exposure outside MCP is prohibited unless explicitly justified for internal system integration.

### III. Test-First Development (NON-NEGOTIABLE)

Tests MUST be written before implementation. The workflow is: Write test → Verify test fails → Implement feature → Verify test passes → Refactor. Red-Green-Refactor cycle is mandatory.

**Rationale**: Test-first design forces clear contracts, prevents regression, and ensures features are independently testable. This is critical for agent integrations where behavior must be deterministic and auditable.

**Enforcement**: Pull requests without accompanying tests that were written first (evidenced by commit history or test failure screenshots) will be rejected. No exceptions.

### IV. Semantic Kernel Agent Integration

Agent orchestration MUST use Semantic Kernel patterns with clear separation between kernel logic, skills/plugins, and model adapters. Skills MUST be small, testable functions with explicit input/output schemas and documented side effects.

**Rationale**: Semantic Kernel provides structured agent development with testable skills, reproducible behavior, and abstraction over local LLM runtimes. This enables swapping models and testing agent logic independently.

**Enforcement**: Agent features MUST NOT embed LLM-specific code in business logic. All model interactions MUST go through Semantic Kernel abstractions. Skills MUST have unit tests with mocked kernel responses.

### V. Self-Contained Deployment

Every component (server, client, chat web UI) MUST publish as a self-contained single-file executable per target platform (Windows, Linux, macOS). Dependencies MUST be bundled; users MUST NOT be required to install runtimes separately.

**Rationale**: Reduces friction for non-technical users and ensures consistent behavior across environments. Self-contained binaries eliminate "works on my machine" issues.

**Enforcement**: Packaging scripts MUST produce single-file executables for win-x64, linux-x64, and osx-x64. Published artifacts MUST include launch wrappers that prompt for required paths and start the executable.

### VI. Structured Observability

All agent decisions, tool invocations, prompt templates used, and model interactions MUST be logged with structured data (JSON or key-value pairs). Logs MUST include correlation IDs for tracing multi-step agent workflows. Telemetry MUST be opt-in and encrypt batched logs before any upload.

**Rationale**: Local agent debugging requires visibility into decision trees and tool chains. Structured logs enable post-mortem analysis and regression detection without requiring live debugging.

**Enforcement**: Every Semantic Kernel skill invocation MUST log: skill name, input parameters, output result, execution time, and correlation ID. Logs MUST be written to a local file or stdout in JSON Lines format.

### VII. Incremental User Value

Features MUST be designed as independently deliverable user stories with clear acceptance criteria. Each story MUST be testable and valuable on its own, enabling MVP-first delivery and parallel development.

**Rationale**: Incremental delivery reduces risk, enables early feedback, and allows teams to validate assumptions before investing in full feature sets.

**Enforcement**: Specifications MUST include prioritized user stories (P1, P2, P3). Implementation plans MUST show how each story can be completed and deployed independently. Tasks MUST be grouped by story with clear checkpoints.

## Technology Stack & Architecture Constraints

**Runtime**: .NET 9.0+ with C# 13 language features. Multi-target libraries when necessary for platform-specific functionality (e.g., secure storage).

**Agent Framework**: Semantic Kernel for agent orchestration, skill/plugin management, and local LLM abstraction. Avoid direct LLM SDK usage in business logic.

**Excel Processing**: ClosedXML library for reading/writing `.xlsx` files. MUST NOT require Microsoft Excel installation or COM automation.

**Local LLM Integration**: Support LM Studio, Ollama, or any OpenAI-compatible local endpoint. Default to `http://localhost:1234` with configurable base URL and model name.

**MCP Transport**: Stdio JSON-RPC for server. Optional HTTP/WebSocket transports for future relay scenarios but stdio MUST remain primary.

**Storage**: SQLite for agent state, conversation history, or indexed metadata. File system for workbooks and logs. Platform secure storage (Credential Locker/Keychain/Keystore) for any API keys or tokens.

**Testing**: xUnit for unit and integration tests. Contract tests for MCP tool/resource schemas. UI automation for chat web app smoke tests.

**Packaging**: Single-file publish with ReadyToRun compilation where supported. Include launch scripts (PowerShell, Bash, Batch) in distribution packages.

## Security & Privacy Requirements

**Secrets Management**: API keys, tokens, or credentials MUST be stored in platform secure storage (Windows Credential Locker, macOS Keychain, Linux Secret Service). Never log secrets or include them in error messages.

**Workbook Access Control**: The server MUST validate that requested workbook paths exist and are readable. File system permissions MUST be respected. No directory traversal attacks allowed.

**Agent Execution Safety**: Semantic Kernel skills that execute code or shell commands MUST run in sandboxed, time-limited contexts with explicit resource caps (memory, CPU). Document which skills perform sensitive operations.

**Data Retention**: Conversation history and cached embeddings MUST be stored locally. Users MUST be able to delete all traces of a session. No persistent identifiers that enable cross-session tracking.

**Model Integrity**: When loading local LLM models, verify checksums or signatures where available. Log model version and source for audit trails.

**Opt-In Telemetry**: Any telemetry upload MUST be explicitly opt-in with clear disclosure of what data is collected. Telemetry MUST be encrypted in transit and at rest before upload.

## Governance

**Constitution Authority**: This constitution supersedes all other development practices, coding standards, or team preferences. When conflicts arise, constitution principles take precedence.

**Amendment Process**: Amendments require:
1. Written proposal documenting the change, rationale, and impact on existing features
2. Review by at least two maintainers
3. Migration plan for any affected code or documentation
4. Version bump according to semantic versioning rules (MAJOR for incompatible changes, MINOR for additions, PATCH for clarifications)

**Compliance Review**: Every pull request MUST include a constitution compliance checklist. Reviewers MUST verify:
- Tests written first (red-green-refactor evidence)
- MCP protocol compliance for agent-facing features
- Local-first operation preserved (no external service dependencies)
- Structured logging added for new agent skills
- Self-contained deployment impact assessed

**Complexity Justification**: Deviations from simplicity principles (e.g., introducing new abstractions, adding architectural layers, creating new projects) MUST be explicitly justified in the implementation plan with evidence that simpler alternatives were considered and rejected.

**Best Practices Reference**: Runtime development guidance for .NET local projects is documented in `docs/best-practices/local-net-bp.md`. Semantic Kernel agent patterns are documented in `docs/best-practices/semantic-kernal-bp.md`. These files provide tactical implementation guidance that operationalizes constitution principles.

**Version**: 1.0.0 | **Ratified**: TODO(RATIFICATION_DATE): original adoption date unknown | **Last Amended**: 2025-10-22
