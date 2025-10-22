# Implementation Plan: Local Excel Conversational Agent

**Branch**: `001-local-excel-chat-agent` | **Date**: October 22, 2025 | **Spec**: [spec.md](./spec.md)
**Input**: Feature specification from `/specs/001-local-excel-chat-agent/spec.md`

**Note**: This template is filled in by the `/speckit.plan` command. See `.specify/templates/commands/plan.md` for the execution workflow.

## Summary

Build a conversational AI agent that enables users to interact with Excel workbooks through natural language queries using a local small language model. The solution extends the existing ExcelMcp.ChatWeb application with Semantic Kernel-based agent orchestration, leveraging the existing MCP server infrastructure for workbook operations. The agent will maintain conversation context (20-turn rolling window), display data as formatted HTML tables, handle multi-workbook sessions, and operate entirely offline on the user's local device. The implementation uses Blazor for the UI, Semantic Kernel for agent logic, and the existing ClosedXML/MCP infrastructure for Excel operations.

## Technical Context

**Language/Version**: C# 13 / .NET 9.0  
**Primary Dependencies**: 
- Microsoft.SemanticKernel (latest stable for .NET 9.0)
- ClosedXML (existing, for Excel file parsing)
- Existing ExcelMcp.Server (MCP protocol implementation)
- Existing ExcelMcp.Contracts (shared data models)

**UI Framework**: Blazor Server (for interactive web UI with real-time updates)  
**Agent Framework**: Semantic Kernel with skills/plugins for MCP tool invocation  
**Storage**: 
- In-memory for conversation context (20-turn rolling window)
- File system for workbooks (existing)
- Local file logging for troubleshooting (JSON Lines format)

**Testing**: xUnit for unit/integration tests, contract tests for MCP compatibility  
**Target Platform**: Windows, Linux, macOS (self-contained single-file executables)  
**Project Type**: Web application with backend agent orchestration  
**Performance Goals**: 
- Workbook load <5 seconds (up to 50MB)
- Query response <10 seconds for structure queries
- Query response <30 seconds for complex data operations
- Support up to 100 sheets, 100k rows

**Constraints**: 
- Must operate entirely offline (no external API calls)
- Must work with small local LLMs (<10GB, consumer hardware)
- No GPU required (optional for performance)
- Must sanitize error messages to protect sensitive data
- Must maintain privacy (no telemetry/data transmission)

**Scale/Scope**: 
- Single-user local application
- Session-based (no persistent multi-user state)
- Support typical business workbooks (dozens of sheets, thousands of rows)

## Constitution Check

*GATE: Must pass before Phase 0 research. Re-check after Phase 1 design.*

| Principle | Status | Evidence/Notes |
|-----------|--------|----------------|
| **I. Privacy-First Local Operation** | ✅ PASS | All processing local; no external API calls (FR-011, FR-012). Workbook data never transmitted. LLM runs locally via LM Studio/Ollama. |
| **II. MCP Protocol Compliance** | ✅ PASS | Leverages existing ExcelMcp.Server with MCP tools. Agent invokes MCP tools via Semantic Kernel plugins. Structured JSON inputs/outputs. |
| **III. Test-First Development** | ⚠️ MONITORED | Must follow red-green-refactor. Plan includes test strategy. Will enforce in PR reviews with commit history checks. |
| **IV. Semantic Kernel Agent Integration** | ✅ PASS | Agent uses Semantic Kernel for orchestration. MCP tools wrapped as SK plugins. Clear separation: kernel logic, skills, model adapters. |
| **V. Self-Contained Deployment** | ✅ PASS | Extends existing packaging scripts. Single-file executables for win-x64, linux-x64, osx-x64. Blazor assets bundled. |
| **VI. Structured Observability** | ✅ PASS | JSON Lines logging for agent decisions, tool invocations, errors. Correlation IDs for multi-turn context. Logs stored locally (FR-010a). |
| **VII. Incremental User Value** | ✅ PASS | Five prioritized user stories (P1-P4). Each independently testable. P1 (Basic Querying) is minimal viable feature. |

**Gate Decision**: ✅ **PROCEED TO PHASE 0**

All critical gates pass. Test-First principle will be enforced through development workflow and PR reviews.

## Project Structure

### Documentation (this feature)

```text
specs/001-local-excel-chat-agent/
├── plan.md              # This file (/speckit.plan command output)
├── research.md          # Phase 0 output (/speckit.plan command)
├── data-model.md        # Phase 1 output (/speckit.plan command)
├── quickstart.md        # Phase 1 output (/speckit.plan command)
├── contracts/           # Phase 1 output (/speckit.plan command)
│   ├── semantic-kernel-plugins.md
│   ├── conversation-api.md
│   └── mcp-tool-mapping.md
├── checklists/          # Existing from /speckit.specify
│   └── requirements.md
└── tasks.md             # Phase 2 output (/speckit.tasks command - NOT created by /speckit.plan)
```

### Source Code (repository root)

```text
src/
├── ExcelMcp.Server/           # EXISTING - MCP server with tools
│   ├── Excel/
│   │   ├── ExcelResourceUri.cs
│   │   └── ExcelWorkbookService.cs
│   └── Mcp/
│       ├── McpServer.cs
│       └── McpModels.cs
│
├── ExcelMcp.Contracts/        # EXISTING - Shared models
│   ├── WorkbookMetadata.cs
│   ├── ResourceContracts.cs
│   └── ExcelSearchContracts.cs
│
├── ExcelMcp.ChatWeb/          # ENHANCED - Add Blazor + Semantic Kernel agent
│   ├── Program.cs             # MODIFY - Add Semantic Kernel services
│   ├── appsettings.json       # MODIFY - Add SK config, logging
│   │
│   ├── Components/            # NEW - Blazor components
│   │   ├── Pages/
│   │   │   ├── Chat.razor           # Main chat interface
│   │   │   └── Chat.razor.cs        # Code-behind with agent interaction
│   │   ├── Shared/
│   │   │   ├── ChatMessage.razor    # Message display component
│   │   │   ├── DataTable.razor      # HTML table renderer
│   │   │   ├── WorkbookSelector.razor
│   │   │   └── LoadingIndicator.razor
│   │   └── Layout/
│   │       └── MainLayout.razor
│   │
│   ├── Services/              # NEW/ENHANCED - Agent and business logic
│   │   ├── Agent/
│   │   │   ├── ExcelAgentService.cs      # Main SK agent orchestrator
│   │   │   ├── ConversationManager.cs    # 20-turn rolling window
│   │   │   ├── PromptTemplates.cs        # SK prompt templates
│   │   │   └── ResponseFormatter.cs      # HTML table generation
│   │   ├── Plugins/           # NEW - SK plugins wrapping MCP
│   │   │   ├── WorkbookStructurePlugin.cs
│   │   │   ├── WorkbookSearchPlugin.cs
│   │   │   ├── DataRetrievalPlugin.cs
│   │   │   └── PluginDescriptors.cs
│   │   ├── McpClientHost.cs   # EXISTING - MCP client
│   │   └── LlmStudioClient.cs # EXISTING - Local LLM client
│   │
│   ├── Models/                # ENHANCED - Add agent models
│   │   ├── ChatDtos.cs        # EXISTING
│   │   ├── ConversationTurn.cs         # NEW
│   │   ├── WorkbookContext.cs          # NEW
│   │   └── AgentResponse.cs            # NEW
│   │
│   ├── Options/               # ENHANCED
│   │   ├── ExcelMcpOptions.cs          # EXISTING
│   │   ├── LlmStudioOptions.cs         # EXISTING
│   │   ├── SemanticKernelOptions.cs    # NEW
│   │   └── ConversationOptions.cs      # NEW (rolling window config)
│   │
│   └── Logging/               # NEW - Structured logging
│       ├── AgentLogger.cs
│       └── CorrelationIdMiddleware.cs
│
└── ExcelMcp.Client/           # EXISTING - CLI client (unchanged)

tests/
├── ExcelMcp.ChatWeb.Tests/   # ENHANCED - Add agent tests
│   ├── Agent/
│   │   ├── ExcelAgentServiceTests.cs
│   │   ├── ConversationManagerTests.cs
│   │   └── ResponseFormatterTests.cs
│   ├── Plugins/
│   │   ├── WorkbookStructurePluginTests.cs
│   │   ├── WorkbookSearchPluginTests.cs
│   │   └── DataRetrievalPluginTests.cs
│   ├── Components/
│   │   └── ChatMessageTests.cs
│   └── Integration/
│       └── EndToEndAgentTests.cs
```

**Structure Decision**: Extend existing `ExcelMcp.ChatWeb` ASP.NET project with Blazor components and Semantic Kernel agent services. This minimizes new projects (adheres to Constitution complexity constraints) while clearly separating concerns: Blazor UI in `/Components`, SK agent logic in `/Services/Agent`, MCP integration via SK plugins in `/Services/Plugins`. Existing ExcelMcp.Server and Contracts remain unchanged.

## Complexity Tracking

> **Constitution compliant - no violations to justify**

This feature extends the existing `ExcelMcp.ChatWeb` project rather than creating new projects. Semantic Kernel integration adds agent orchestration capabilities within established boundaries. No additional architectural layers or abstractions beyond what Semantic Kernel provides. The design follows existing patterns: Services for business logic, Models for data, Options for configuration.

---

## Phase 0: Research & Technology Decisions ✅ COMPLETE

**Artifacts Created**:
- `research.md` - Technology decisions, patterns, and risk mitigation

**Key Decisions**:
1. **SK Plugin Architecture** - Wrap MCP tools as Semantic Kernel plugins
2. **Rolling Window** - In-memory circular buffer with SK ChatHistory (20 turns)
3. **HTML Tables** - Blazor Razor components for data rendering
4. **Session State** - Blazor scoped services for multi-workbook tracking
5. **Two-Tier Logging** - Sanitized UI + detailed JSON Lines logs

**Dependencies Identified**:
- Microsoft.SemanticKernel (latest stable for .NET 9.0)
- Serilog.AspNetCore + Serilog.Sinks.File + Serilog.Formatting.Compact
- Existing: ClosedXML, ExcelMcp.Server, ExcelMcp.Contracts

---

## Phase 1: Design & Contracts ✅ COMPLETE

**Artifacts Created**:
- `data-model.md` - Entity definitions with validation rules
- `contracts/semantic-kernel-plugins.md` - SK plugin method specifications
- `contracts/conversation-api.md` - Service interface contracts
- `contracts/mcp-tool-mapping.md` - MCP to SK plugin mappings
- `quickstart.md` - Developer getting-started guide
- `.github/copilot-instructions.md` - Updated with C# 13 / .NET 9.0 context

**Data Models Defined**:
- ConversationTurn, WorkbookContext, WorkbookSession
- AgentResponse, ToolInvocation, SanitizedError
- ContentType enum, ErrorCode enum

**API Contracts Defined**:
- IExcelAgentService (ProcessQueryAsync, LoadWorkbookAsync, etc.)
- IConversationManager (AddUserTurn, GetContextForLLM, etc.)
- IResponseFormatter (FormatAsHtmlTable, SanitizeErrorMessage, etc.)

**SK Plugins Specified**:
- WorkbookStructurePlugin (3 methods)
- WorkbookSearchPlugin (2 methods)
- DataRetrievalPlugin (3 methods)

**Constitution Re-Check**: ✅ All principles satisfied post-design

---

## Phase 2: Implementation Planning (Next Step: `/speckit.tasks`)

**Status**: Ready for task breakdown

**Recommended Task Organization**:

### Epic 1: Core Infrastructure (P1 - Week 1)
- Data models implementation
- Service interfaces
- Configuration options
- Logging infrastructure

### Epic 2: MCP Integration (P1 - Week 2)
- Semantic Kernel plugin implementations
- MCP client enhancements (timeout, correlation IDs)
- Error handling and sanitization
- Plugin unit tests

### Epic 3: Agent Orchestration (P1 - Week 2-3)
- ConversationManager implementation
- ExcelAgentService with SK integration
- ResponseFormatter with HTML table generation
- Prompt templates
- Integration tests

### Epic 4: Blazor UI (P1 - Week 3-4)
- Chat page component
- Message display components
- Workbook selector
- Loading indicators
- CSS styling
- Component tests

### Epic 5: Multi-Workbook Support (P2 - Week 4)
- Session state enhancements
- Workbook switch markers
- Context preservation
- UI indicators

### Epic 6: Advanced Queries (P3/P4 - Week 5+)
- Cross-sheet analysis
- Data filtering
- Export functionality
- Additional SK plugins as needed

**Testing Strategy**:
- Unit tests for each model, service, plugin
- Integration tests for agent workflows
- Contract tests for MCP/SK interfaces
- End-to-end tests with sample workbooks

**Performance Targets** (from Success Criteria):
- Workbook load: <5s (50MB files)
- Structure queries: <10s
- Complex queries: <30s
- Support: 100 sheets, 100k rows

---

## Next Command

Run `/speckit.tasks` to generate detailed task breakdown with estimates, dependencies, and test scenarios.
