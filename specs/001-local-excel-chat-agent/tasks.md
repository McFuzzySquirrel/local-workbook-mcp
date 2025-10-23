# Tasks: Local Excel Conversational Agent

**Feature Branch**: `001-local-excel-chat-agent`  
**Input**: Design documents from `/specs/001-local-excel-chat-agent/`  
**Prerequisites**: plan.md, spec.md, data-model.md, contracts/

**Organization**: Tasks are grouped by user story to enable independent implementation and testing of each story.

---

## Format: `[ID] [P?] [Story] Description`

- **[P]**: Can run in parallel (different files, no dependencies)
- **[Story]**: Which user story this task belongs to (e.g., US1, US2, US3, US4, US5)
- Include exact file paths in descriptions

**Path Conventions**: All paths relative to repository root `d:\GitHub Projects\spec-kit-demos\local-workbook-mcp\`

---

## Phase 1: Setup (Shared Infrastructure)

**Purpose**: Project initialization and basic structure

- [X] T001 Add Semantic Kernel NuGet package to `src/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb.csproj` (Microsoft.SemanticKernel latest stable)
- [X] T002 [P] Add Serilog packages to `src/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb.csproj` (Serilog.AspNetCore, Serilog.Sinks.File, Serilog.Formatting.Compact)
- [X] T003 [P] Create directory structure: `src/ExcelMcp.ChatWeb/Components/Pages/`, `Components/Shared/`, `Components/Layout/`
- [X] T004 [P] Create directory structure: `src/ExcelMcp.ChatWeb/Services/Agent/`, `Services/Plugins/`, `Logging/`
- [X] T005 Update `src/ExcelMcp.ChatWeb/appsettings.json` with SemanticKernel options (model, baseUrl, timeout)
- [X] T006 Update `src/ExcelMcp.ChatWeb/appsettings.json` with Conversation options (maxContextTurns: 20)
- [X] T007 Update `src/ExcelMcp.ChatWeb/appsettings.json` with Serilog JSON Lines logging configuration (logs/agent-{Date}.log)

**Checkpoint**: ✅ Dependencies installed, directory structure ready, configuration in place

---

## Phase 2: Foundational (Blocking Prerequisites)

**Purpose**: Core infrastructure that MUST be complete before ANY user story can be implemented

**⚠️ CRITICAL**: No user story work can begin until this phase is complete

### Data Models

- [X] T008 [P] Create `src/ExcelMcp.ChatWeb/Models/ConversationTurn.cs` with properties: Id, Role, Content, Timestamp, CorrelationId, ContentType, Metadata
- [X] T009 [P] Create `src/ExcelMcp.ChatWeb/Models/WorkbookContext.cs` with properties: WorkbookPath, WorkbookName, LoadedAt, Metadata, IsValid, ErrorMessage
- [X] T010 [P] Create `src/ExcelMcp.ChatWeb/Models/WorkbookSession.cs` with properties: SessionId, StartedAt, LastActivityAt, CurrentContext, ConversationHistory, ContextWindow, PreviousContexts
- [X] T011 [P] Create `src/ExcelMcp.ChatWeb/Models/AgentResponse.cs` with properties: ResponseId, CorrelationId, Content, ContentType, TableData, Suggestions, ToolsInvoked, ProcessingTimeMs, ModelUsed, Error
- [X] T012 [P] Create `src/ExcelMcp.ChatWeb/Models/ToolInvocation.cs` with properties: ToolName, PluginMethod, InputParameters, OutputSummary, InvokedAt, DurationMs, Success, ErrorMessage
- [X] T013 [P] Create `src/ExcelMcp.ChatWeb/Models/SanitizedError.cs` with properties: Message, CorrelationId, Timestamp, ErrorCode, CanRetry, SuggestedAction
- [X] T014 [P] Create `src/ExcelMcp.ChatWeb/Models/TableData.cs` with properties: Columns, Rows, Metadata (includes SheetName, RowCount, IsTruncated)
- [X] T015 [P] Create `src/ExcelMcp.ChatWeb/Models/ContentType.cs` enum with values: Text, Table, Error, SystemMessage, Clarification
- [X] T016 [P] Create `src/ExcelMcp.ChatWeb/Models/ErrorCode.cs` enum with values: WorkbookLoadFailed, QueryTimeout, ModelUnresponsive, InvalidQuery, McpToolError, UnknownError

### Configuration Options

- [X] T017 [P] Create `src/ExcelMcp.ChatWeb/Options/SemanticKernelOptions.cs` with properties: Model, BaseUrl, ApiKey, TimeoutSeconds
- [X] T018 [P] Create `src/ExcelMcp.ChatWeb/Options/ConversationOptions.cs` with properties: MaxContextTurns, MaxResponseLength, SuggestedQueriesCount

### Logging Infrastructure

- [X] T019 Create `src/ExcelMcp.ChatWeb/Logging/AgentLogger.cs` with methods: LogQuery, LogToolInvocation, LogResponse, LogError (with correlationId)
- [X] T020 Create `src/ExcelMcp.ChatWeb/Logging/CorrelationIdMiddleware.cs` for HTTP request correlation tracking

### Service Interfaces

- [X] T021 Create `src/ExcelMcp.ChatWeb/Services/Agent/IExcelAgentService.cs` interface with methods: ProcessQueryAsync, LoadWorkbookAsync, ClearConversationAsync, GetSuggestedQueriesAsync
- [X] T022 Create `src/ExcelMcp.ChatWeb/Services/Agent/IConversationManager.cs` interface with methods: AddUserTurn, AddAssistantTurn, AddSystemMessage, GetContextForLLM, GetFullHistory
- [X] T023 Create `src/ExcelMcp.ChatWeb/Services/Agent/IResponseFormatter.cs` interface with methods: FormatAsHtmlTable, FormatAsText, SanitizeErrorMessage

### Program.cs Configuration

- [X] T024 Update `src/ExcelMcp.ChatWeb/Program.cs` to register Semantic Kernel with OpenAI chat completion (using LlmStudioOptions for endpoint)
- [X] T025 Update `src/ExcelMcp.ChatWeb/Program.cs` to configure Serilog with JSON Lines file sink
- [X] T026 Update `src/ExcelMcp.ChatWeb/Program.cs` to register scoped WorkbookSession service
- [X] T027 Update `src/ExcelMcp.ChatWeb/Program.cs` to add CorrelationIdMiddleware to request pipeline

**Checkpoint**: ✅ Foundation ready - all models, interfaces, and core infrastructure complete. User story implementation can now begin in parallel

---

## Phase 3: User Story 1 - Basic Workbook Querying (Priority: P1) 🎯 MVP

**Goal**: Users can load an Excel workbook, ask natural language questions about structure and data, and receive accurate responses. This is the minimal viable product.

**Independent Test**: Load sample workbook, ask "What sheets are in this workbook?", verify agent returns accurate sheet names.

### MCP-Wrapping Plugins (Core Dependencies)

- [X] T028 [P] [US1] Create `src/ExcelMcp.ChatWeb/Services/Plugins/WorkbookStructurePlugin.cs` with constructor accepting IMcpClient
- [X] T029 [US1] Implement `WorkbookStructurePlugin.ListWorkbookStructure()` method with [KernelFunction] attribute, calling MCP excel-list-structure tool
- [X] T030 [US1] Implement `WorkbookStructurePlugin.GetSheetNames()` method, deriving from excel-list-structure response
- [X] T031 [US1] Implement `WorkbookStructurePlugin.GetTableInfo(string sheetName)` method, filtering excel-list-structure by sheet
- [X] T032 [P] [US1] Create `src/ExcelMcp.ChatWeb/Services/Plugins/WorkbookSearchPlugin.cs` with constructor accepting IMcpClient
- [X] T033 [US1] Implement `WorkbookSearchPlugin.SearchWorkbook(string searchText, int maxResults)` method calling MCP excel-search tool
- [X] T034 [US1] Implement `WorkbookSearchPlugin.SearchInSheet(string sheetName, string searchText, int maxResults)` method with sheet filter
- [X] T035 [P] [US1] Create `src/ExcelMcp.ChatWeb/Services/Plugins/DataRetrievalPlugin.cs` with constructor accepting IMcpClient
- [X] T036 [US1] Implement `DataRetrievalPlugin.PreviewTable(string name, int rowCount, int startRow)` method calling MCP excel-preview-table tool
- [X] T037 [US1] Implement `DataRetrievalPlugin.GetRowsInRange(string sheetName, string cellRange)` method using MCP excel-preview-table with range
- [X] T038 [US1] Implement `DataRetrievalPlugin.CalculateAggregation(string name, string column, string aggregationType)` method (retrieve data + calculate)
- [X] T039 [US1] Add error handling to all plugin methods to return standardized error JSON with codes: NO_WORKBOOK, SHEET_NOT_FOUND, MCP_ERROR, etc.

### Response Formatting

- [X] T040 [US1] Create `src/ExcelMcp.ChatWeb/Services/Agent/ResponseFormatter.cs` implementing IResponseFormatter
- [X] T041 [US1] Implement `ResponseFormatter.FormatAsHtmlTable(TableData)` method generating HTML table with headers, borders, styling
- [X] T042 [US1] Implement `ResponseFormatter.FormatAsText(string content)` method converting markdown-like syntax to HTML
- [X] T043 [US1] Implement `ResponseFormatter.SanitizeErrorMessage(Exception, correlationId)` method mapping exception to SanitizedError with generic messages

### Conversation Management

- [X] T044 [US1] Create `src/ExcelMcp.ChatWeb/Services/Agent/ConversationManager.cs` implementing IConversationManager
- [X] T045 [US1] Implement `ConversationManager.AddUserTurn(message, correlationId)` adding to full history + context window with eviction (20-turn limit)
- [X] T046 [US1] Implement `ConversationManager.AddAssistantTurn(message, correlationId, contentType)` with same rolling window logic
- [X] T047 [US1] Implement `ConversationManager.AddSystemMessage(message)` adding to full history only (NOT context window)
- [X] T048 [US1] Implement `ConversationManager.GetContextForLLM()` returning SK ChatHistory instance with last 20 turns
- [X] T049 [US1] Implement `ConversationManager.GetFullHistory()` returning complete list of ConversationTurn for UI display

### Agent Service (Core Orchestration)

- [X] T050 [US1] Create `src/ExcelMcp.ChatWeb/Services/Agent/ExcelAgentService.cs` implementing IExcelAgentService
- [X] T051 [US1] Inject Kernel, IConversationManager, IResponseFormatter, IMcpClient, ILogger into ExcelAgentService constructor
- [X] T052 [US1] Implement `ExcelAgentService.LoadWorkbookAsync(filePath, session)` method: validate file, start MCP server, retrieve metadata, create WorkbookContext
- [X] T053 [US1] Add system message insertion to LoadWorkbookAsync when workbook changes (e.g., "Workbook changed to Budget.xlsx")
- [X] T054 [US1] Implement `ExcelAgentService.ProcessQueryAsync(query, session)` method: validate query, add user turn, invoke SK with context, format response, add assistant turn
- [X] T055 [US1] Add timeout handling (30 seconds) to ProcessQueryAsync using CancellationToken
- [X] T056 [US1] Add error handling to ProcessQueryAsync: catch exceptions, sanitize errors, return error AgentResponse with retry option
- [X] T057 [US1] Implement `ExcelAgentService.ClearConversationAsync(session)` method clearing history and context window (keep workbook loaded)
- [X] T058 [US1] Implement `ExcelAgentService.GetSuggestedQueriesAsync(session, maxSuggestions)` analyzing workbook metadata and recent turns for suggestions

### Plugin Registration

- [X] T059 [US1] Update `src/ExcelMcp.ChatWeb/Program.cs` to register WorkbookStructurePlugin, WorkbookSearchPlugin, DataRetrievalPlugin as singletons
- [X] T060 [US1] Update `src/ExcelMcp.ChatWeb/Program.cs` to add plugins to Semantic Kernel using Plugins.AddFromObject for each plugin
- [X] T061 [US1] Register IExcelAgentService, IConversationManager, IResponseFormatter in dependency injection container

### Blazor UI Components

- [X] T062 [P] [US1] Create `src/ExcelMcp.ChatWeb/Components/Shared/WorkbookSelector.razor` with file input and load button
- [X] T063 [US1] Add event callback `OnWorkbookSelected` to WorkbookSelector emitting selected file path
- [X] T064 [P] [US1] Create `src/ExcelMcp.ChatWeb/Components/Shared/ChatMessage.razor` component accepting ConversationTurn parameter
- [X] T065 [US1] Add rendering logic to ChatMessage: display role, timestamp, content (text or HTML table based on ContentType)
- [X] T066 [P] [US1] Create `src/ExcelMcp.ChatWeb/Components/Shared/DataTable.razor` component accepting TableData parameter
- [X] T067 [US1] Implement DataTable rendering HTML table with columns from TableData.Columns and rows from TableData.Rows
- [X] T068 [P] [US1] Create `src/ExcelMcp.ChatWeb/Components/Shared/LoadingIndicator.razor` component showing spinner during processing
- [X] T069 [P] [US1] Create `src/ExcelMcp.ChatWeb/Components/Pages/Chat.razor` main chat page with WorkbookSelector, message list, input box
- [X] T070 [US1] Create `src/ExcelMcp.ChatWeb/Components/Pages/Chat.razor.cs` code-behind injecting IExcelAgentService and WorkbookSession
- [X] T071 [US1] Implement `Chat.razor.cs.OnWorkbookSelected(filePath)` handler calling LoadWorkbookAsync and displaying result
- [X] T072 [US1] Implement `Chat.razor.cs.OnQuerySubmitted(query)` handler calling ProcessQueryAsync with loading state and error handling
- [X] T073 [US1] Add CancellationTokenSource to Chat.razor.cs for timeout support, disposing on component disposal
- [X] T074 [US1] Bind message list in Chat.razor to session.ConversationHistory for display
- [X] T075 [US1] Add "Clear History" button to Chat.razor triggering ClearConversationAsync
- [X] T076 [P] [US1] Create `src/ExcelMcp.ChatWeb/Components/Layout/MainLayout.razor` basic layout with title and main content area
- [X] T077 [P] [US1] Add CSS styles to `src/ExcelMcp.ChatWeb/wwwroot/styles.css` for chat messages, tables, loading indicator

### Validation & Integration

- [ ] T078 [US1] Test workbook load flow: select file → verify metadata displayed → confirm "ready for queries" message
- [ ] T079 [US1] Test basic query: load workbook → ask "What sheets are in this workbook?" → verify accurate sheet list response
- [ ] T080 [US1] Test data retrieval: ask "Show me first 10 rows of Sales table" → verify HTML table rendered correctly
- [ ] T081 [US1] Test error handling: ask query with no workbook loaded → verify user-friendly error + prompt to load workbook
- [ ] T082 [US1] Test error handling: load corrupted file → verify sanitized error message + correlationId for troubleshooting
- [ ] T083 [US1] Test query timeout: simulate slow MCP response → verify 30-second timeout → user sees error with retry option
- [ ] T084 [US1] Run quickstart.md validation: follow developer setup guide → verify all User Story 1 scenarios work end-to-end

**Checkpoint**: At this point, User Story 1 should be fully functional and testable independently. Users can load workbooks and query them successfully. This is the MVP.

---

## Phase 4: User Story 2 - Multi-Turn Conversations with Context (Priority: P2)

**Goal**: Enable natural conversation flow where follow-up questions reference previous queries without repeating context.

**Independent Test**: Load workbook, ask "What's in the Sales sheet?", then follow up with "Show me the top 5 rows" without re-specifying sheet name. Verify agent understands context.

### Prompt Engineering

- [ ] T085 [US2] Create `src/ExcelMcp.ChatWeb/Services/Agent/PromptTemplates.cs` with system prompt template emphasizing context awareness
- [ ] T086 [US2] Add prompt template that instructs LLM to reference previous conversation turns for pronoun resolution (it, that, those)
- [ ] T087 [US2] Add example prompts to PromptTemplates demonstrating multi-turn context handling (e.g., Q1: "Sales data?", Q2: "Show top 5")

### Enhanced Agent Logic

- [ ] T088 [US2] Update `ExcelAgentService.ProcessQueryAsync` to include system prompt from PromptTemplates when invoking Semantic Kernel
- [ ] T089 [US2] Update `ExcelAgentService.ProcessQueryAsync` to pass full 20-turn context window to SK for every query
- [ ] T090 [US2] Add conversation summarization to `ExcelAgentService` when user asks "Can you summarize what we've discussed?"

### UI Enhancements

- [ ] T091 [P] [US2] Add conversation summary button to `src/ExcelMcp.ChatWeb/Components/Pages/Chat.razor` triggering summarization query
- [ ] T092 [P] [US2] Display turn count indicator in Chat.razor showing "X/20 turns in context" to help users understand window limits

### Validation

- [ ] T093 [US2] Test multi-turn flow: ask "What tables are in Budget sheet?" → "Show me the first one" → verify agent displays first table without re-asking sheet name
- [ ] T094 [US2] Test pronoun resolution: ask about revenue → follow up "What about last quarter?" → verify agent maintains context
- [ ] T095 [US2] Test conversation summary: conduct 5-turn conversation → request summary → verify recap includes key findings
- [ ] T096 [US2] Test context window eviction: conduct 21+ turns → verify oldest turns evicted from LLM context but still visible in UI history

**Checkpoint**: At this point, User Stories 1 AND 2 should both work independently. Natural conversation flow is functional.

---

## Phase 5: User Story 3 - Cross-Sheet Data Insights (Priority: P2)

**Goal**: Enable queries that analyze data across multiple worksheets or tables, correlating information from different sources.

**Independent Test**: Load workbook with Sales and Returns sheets, ask "Are there any products that appear in both Sales and Returns?", verify agent identifies common entries.

### Enhanced Plugin Logic

- [ ] T097 [US3] Update `WorkbookSearchPlugin.SearchWorkbook` to return sheet-grouped results for easier cross-sheet analysis
- [ ] T098 [US3] Add helper method to `DataRetrievalPlugin` for retrieving data from multiple sheets in single call (e.g., PreviewMultipleSheets)
- [ ] T099 [US3] Add caching to plugin methods to avoid redundant MCP calls when querying same data across multiple sheets

### Prompt Engineering for Cross-Sheet Analysis

- [ ] T100 [US3] Update `PromptTemplates.cs` to include examples of cross-sheet queries and expected tool invocation patterns
- [ ] T101 [US3] Add guidance in system prompt for LLM to identify when multiple sheets need correlation (e.g., matching product IDs)

### Validation

- [ ] T102 [US3] Test cross-sheet search: ask "Find all mentions of Project X across entire workbook" → verify results from multiple sheets
- [ ] T103 [US3] Test cross-sheet correlation: ask "Which products have low inventory but high sales?" → verify agent queries both Inventory and Sales sheets
- [ ] T104 [US3] Test multi-sheet aggregation: ask "Compare revenue trends across all quarterly sheets" → verify agent aggregates data from Q1, Q2, Q3, Q4 sheets

**Checkpoint**: All user stories 1-3 should now be independently functional. Cross-sheet analysis capabilities working.

---

## Phase 6: User Story 4 - Data Filtering and Searching (Priority: P3)

**Goal**: Support targeted queries with criteria-based filtering (value ranges, text matching, date filters).

**Independent Test**: Load sales workbook, ask "Show me all sales over $10,000 in January", verify only matching rows returned.

### Enhanced Data Retrieval

- [ ] T105 [US4] Add filtering logic to `DataRetrievalPlugin.PreviewTable` to support WHERE-clause-like conditions (value range, text match, date range)
- [ ] T106 [US4] Implement numeric range filtering in DataRetrievalPlugin (e.g., "sales between $5,000 and $15,000")
- [ ] T107 [US4] Implement text pattern matching in DataRetrievalPlugin (e.g., "customers in California")
- [ ] T108 [US4] Implement multi-criteria filtering in DataRetrievalPlugin (e.g., "Q3 sales over $20,000 from West region")

### Prompt Guidance

- [ ] T109 [US4] Update `PromptTemplates.cs` with examples of filter-based queries and how to invoke filtering parameters
- [ ] T110 [US4] Add error handling for ambiguous filter criteria (e.g., "high sales" without numeric threshold) → agent asks clarifying question

### Validation

- [ ] T111 [US4] Test numeric range filter: ask "Show sales between $5,000 and $15,000" → verify only matching rows returned
- [ ] T112 [US4] Test text matching: ask "Find all entries for customers in California" → verify filtered results
- [ ] T113 [US4] Test multi-criteria: ask "Show Q3 sales over $20,000 from West region" → verify all filters applied correctly
- [ ] T114 [US4] Test ambiguous filter: ask "Show high sales" → verify agent requests clarification on threshold

**Checkpoint**: All user stories 1-4 independently functional. Advanced filtering capabilities working.

---

## Phase 7: User Story 5 - Export Conversation Insights (Priority: P4)

**Goal**: Allow users to save or share insights, conversation history, or specific data views.

**Independent Test**: Conduct query session, request "Export our conversation", verify readable summary file created.

### Export Service

- [ ] T115 [P] [US5] Create `src/ExcelMcp.ChatWeb/Services/ExportService.cs` with methods: ExportConversation, ExportDataView, ExportInsightsSummary
- [ ] T116 [US5] Implement `ExportService.ExportConversation(session)` generating Markdown file with full conversation history (user queries + agent responses)
- [ ] T117 [US5] Implement `ExportService.ExportDataView(TableData)` generating CSV file from table data
- [ ] T118 [US5] Implement `ExportService.ExportInsightsSummary(session)` using SK to generate concise summary of key findings from conversation
- [ ] T119 [US5] Add file download endpoint to `src/ExcelMcp.ChatWeb/Program.cs` for serving generated export files

### UI Components

- [ ] T120 [P] [US5] Add "Export Conversation" button to `src/ExcelMcp.ChatWeb/Components/Pages/Chat.razor` triggering ExportConversation
- [ ] T121 [P] [US5] Add "Export Table" button to `src/ExcelMcp.ChatWeb/Components/Shared/DataTable.razor` when table displayed
- [ ] T122 [US5] Add "Generate Summary" button to Chat.razor triggering ExportInsightsSummary with LLM-generated report

### Validation

- [ ] T123 [US5] Test conversation export: complete 5-turn conversation → click "Export Conversation" → verify Markdown file downloaded with all turns
- [ ] T124 [US5] Test data view export: display table → click "Export Table" → verify CSV file contains correct columns and rows
- [ ] T125 [US5] Test insights summary: conduct analysis session → click "Generate Summary" → verify LLM-generated concise report with key findings

**Checkpoint**: All user stories 1-5 independently functional. Export and sharing capabilities complete.

---

## Phase 8: Polish & Cross-Cutting Concerns

**Purpose**: Improvements that affect multiple user stories

### Documentation

- [ ] T126 [P] Update `README.md` with feature overview, setup instructions, and link to quickstart.md
- [ ] T127 [P] Create `docs/UserGuide.md` with end-user instructions for using the chat interface
- [ ] T128 [P] Update `docs/FutureFeatures.md` with potential enhancements beyond current scope

### Code Quality

- [ ] T129 Code cleanup: remove unused usings, apply consistent formatting across all new files
- [ ] T130 Refactoring: extract common error handling patterns into shared utility class
- [ ] T131 Add XML documentation comments to all public APIs in Services/Agent/ and Services/Plugins/

### Performance Optimization

- [ ] T132 Profile workbook load time → optimize if exceeding 5-second target for 50MB files
- [ ] T133 Profile query response time → optimize if exceeding 10-second target for structure queries
- [ ] T134 Add response caching for repeated identical queries within same session

### Security Hardening

- [ ] T135 Audit error messages across all components to ensure no sensitive data exposure (file paths, sheet names, cell values)
- [ ] T136 Add input validation for file paths in WorkbookSelector to prevent path traversal attacks
- [ ] T137 Review all AgentLogger calls to ensure sanitization before logging user queries containing potential PII

### Final Validation

- [ ] T138 Run complete quickstart.md validation end-to-end following developer guide
- [ ] T139 Test on sample workbooks from `examples/` directory covering all user scenarios
- [ ] T140 Verify all 10 success criteria from spec.md (SC-001 through SC-010) are met
- [ ] T141 Perform constitution compliance check: Privacy-First, MCP Protocol, SK Integration, Self-Contained Deployment, Structured Observability

**Checkpoint**: Feature complete, polished, and ready for release

---

## Dependencies & Execution Order

### Phase Dependencies

- **Setup (Phase 1)**: No dependencies - can start immediately
- **Foundational (Phase 2)**: Depends on Setup completion - BLOCKS all user stories
- **User Story 1 (Phase 3)**: Depends on Foundational phase completion - MVP critical path
- **User Story 2 (Phase 4)**: Depends on User Story 1 completion (extends conversation capabilities)
- **User Story 3 (Phase 5)**: Depends on User Story 1 completion (can run in parallel with US2)
- **User Story 4 (Phase 6)**: Depends on User Story 1 completion (can run in parallel with US2/US3)
- **User Story 5 (Phase 7)**: Depends on User Story 1 completion (can run in parallel with US2/US3/US4)
- **Polish (Phase 8)**: Depends on all desired user stories being complete

### User Story Dependencies

- **User Story 1 (P1 - MVP)**: Can start after Foundational (Phase 2) - No dependencies on other stories
- **User Story 2 (P2)**: Extends US1 conversation features - Requires US1 core components but is independently testable
- **User Story 3 (P2)**: Extends US1 data analysis - Requires US1 plugins but is independently testable, can run parallel with US2
- **User Story 4 (P3)**: Extends US1 data retrieval - Requires US1 plugins but is independently testable, can run parallel with US2/US3
- **User Story 5 (P4)**: Extends US1 with export - Requires US1 conversation manager but is independently testable, can run parallel with others

### Within Each User Story

**User Story 1 Critical Path**:
1. Models (T008-T016) must complete before services can be implemented
2. Interfaces (T021-T023) must exist before implementations
3. Plugins (T028-T039) must exist before agent service can invoke them
4. ResponseFormatter (T040-T043) must exist before agent service can format responses
5. ConversationManager (T044-T049) must exist before agent service can manage history
6. ExcelAgentService (T050-T058) must complete before UI components can use it
7. UI Components (T062-T077) can mostly run in parallel after agent service ready
8. Validation (T078-T084) runs after all implementation complete

**Other User Stories**:
- Build incrementally on User Story 1 foundation
- Enhance existing components rather than creating new ones
- Each story should be completable and testable independently

### Parallel Opportunities

**Phase 1 - Setup**: All tasks (T001-T007) can run in parallel (different files, independent operations)

**Phase 2 - Foundational**:
- All Models (T008-T016) can run in parallel
- All Options (T017-T018) can run in parallel
- Logging tasks (T019-T020) can run in parallel
- Interface definitions (T021-T023) can run in parallel
- Program.cs updates (T024-T027) must be sequential (same file)

**Phase 3 - User Story 1**:
- Plugin creation (T028, T032, T035) can run in parallel (different files)
- Plugin method implementations within same plugin must be sequential (same file)
- ResponseFormatter and ConversationManager implementations can run in parallel (different files)
- UI component creation (T062, T064, T066, T068, T069, T076) can run in parallel (different razor files)

**Phase 4-7 - User Stories 2-5**:
- After User Story 1 complete, all remaining user stories can run in parallel (if team capacity allows)
- Each story touches different methods/components
- US2 touches prompts, US3 touches plugins, US4 touches filtering, US5 adds export service

**Phase 8 - Polish**:
- Documentation tasks (T126-T128) can run in parallel
- Performance tasks (T132-T134) can run in parallel
- Security audits (T135-T137) can run in parallel
- Final validation (T138-T141) should run sequentially

---

## Parallel Example: User Story 1 Foundation

```bash
# Launch all model files in parallel (different files):
Task T008: Create ConversationTurn.cs
Task T009: Create WorkbookContext.cs
Task T010: Create WorkbookSession.cs
Task T011: Create AgentResponse.cs
Task T012: Create ToolInvocation.cs
Task T013: Create SanitizedError.cs
Task T014: Create TableData.cs
Task T015: Create ContentType.cs enum
Task T016: Create ErrorCode.cs enum

# After models complete, launch plugin files in parallel:
Task T028: Create WorkbookStructurePlugin.cs
Task T032: Create WorkbookSearchPlugin.cs
Task T035: Create DataRetrievalPlugin.cs

# Also in parallel (different files):
Task T040: Create ResponseFormatter.cs
Task T044: Create ConversationManager.cs
```

---

## Parallel Example: User Stories 2-5 After US1

```bash
# After User Story 1 (MVP) complete and validated, launch in parallel:

Team Member A - User Story 2 (Multi-Turn Context):
  Task T085-T092: Prompt engineering and context management

Team Member B - User Story 3 (Cross-Sheet Insights):
  Task T097-T104: Enhanced plugins for cross-sheet analysis

Team Member C - User Story 4 (Data Filtering):
  Task T105-T114: Filtering logic and validation

Team Member D - User Story 5 (Export):
  Task T115-T125: Export service and UI components
```

---

## Implementation Strategy

### MVP First (Recommended for Single Developer)

1. ✅ Complete Phase 1: Setup (install packages, create directories, configure settings)
2. ✅ Complete Phase 2: Foundational (CRITICAL - all models, interfaces, base services)
3. ✅ Complete Phase 3: User Story 1 (Basic Workbook Querying - full MVP)
4. 🎯 **STOP and VALIDATE**: Test User Story 1 independently with sample workbooks
5. 📦 **Deploy/Demo MVP** if ready (this is shippable!)
6. Proceed to Phase 4+ based on priority

### Incremental Delivery (Recommended)

1. Setup + Foundational → Foundation ready ✅
2. Add User Story 1 → Test independently → Deploy/Demo (MVP!) 🎯
3. Add User Story 2 → Test independently → Deploy/Demo (better conversations)
4. Add User Story 3 → Test independently → Deploy/Demo (cross-sheet insights)
5. Add User Story 4 → Test independently → Deploy/Demo (advanced filtering)
6. Add User Story 5 → Test independently → Deploy/Demo (export capabilities)
7. Polish (Phase 8) → Final release

Each story adds value without breaking previous stories!

### Parallel Team Strategy (If Multiple Developers)

With multiple developers:

1. **Week 1**: Team completes Setup + Foundational together
2. **Week 2-3**: Once Foundational done:
   - Lead Developer: User Story 1 (MVP critical path)
   - Other developers: Prepare tests, documentation, infrastructure
3. **Week 4**: After User Story 1 validated:
   - Developer A: User Story 2 (Multi-Turn Context)
   - Developer B: User Story 3 (Cross-Sheet Insights)
   - Developer C: User Story 4 (Data Filtering)
   - Developer D: User Story 5 (Export)
4. **Week 5**: Integration testing, polish, final validation

Stories complete and integrate independently!

---

## Risk Mitigation

### High-Risk Tasks (Need Extra Attention)

- **T054-T056**: `ProcessQueryAsync` implementation - Core agent logic, many failure modes, complex error handling
- **T029-T038**: Plugin implementations - Must correctly wrap MCP tools, handle all error cases
- **T041**: HTML table formatting - Security risk (XSS) if not properly escaping cell values
- **T043**: Error sanitization - Privacy risk if sensitive data leaks into error messages
- **T072**: Query submission handler - Must handle timeouts, cancellation, errors gracefully

### Critical Validation Points

- After T027: Verify all DI registrations correct before proceeding to user stories
- After T039: Verify all plugins working with MCP server using simple test queries
- After T058: Verify agent service can successfully process query end-to-end
- After T077: Verify complete UI rendering with real workbook data
- After T084: Verify all User Story 1 acceptance scenarios pass (this is MVP gate!)

### Fallback Plans

- If Semantic Kernel integration proves problematic → Simplify to direct OpenAI API calls with manual tool routing
- If LLM function calling unreliable → Fall back to prompt-based tool selection with regex parsing
- If HTML table rendering performance poor → Switch to virtualized/paginated table component
- If 20-turn context window insufficient → Make configurable in ConversationOptions

---

## Notes

- **[P] tasks** = different files, no dependencies, can run in parallel
- **[Story] label** maps task to specific user story for traceability
- Each user story should be **independently completable and testable**
- **Test-first development**: For production code, write tests before implementation (not included in this task list per spec)
- **Commit strategy**: Commit after each completed task or logical group of related tasks
- **Stop at checkpoints** to validate story independently before proceeding
- **Avoid**: vague tasks, same-file conflicts, cross-story dependencies that break independence
- **Constitution compliance**: Verify Privacy-First (no data transmission), MCP Protocol (use existing tools), SK Integration (proper plugin patterns), Self-Contained Deployment (all local)

---

## Success Metrics (From spec.md)

Track progress against these success criteria:

- **SC-001**: Workbook load < 5 seconds for 50MB files
- **SC-002**: Basic queries < 10 seconds response time
- **SC-003**: 90% accuracy on structure queries
- **SC-004**: 95% accuracy on data retrieval
- **SC-005**: Complete workflow in < 3 minutes
- **SC-006**: Handle 100 sheets, 100k rows without crashing
- **SC-007**: 80% users complete first query without help
- **SC-008**: Context maintained across 5+ consecutive queries
- **SC-009**: 100% of common failures have helpful error messages
- **SC-010**: Runs on 8GB RAM, standard CPU, no GPU required

Validate these metrics after completing User Story 1 (MVP) and again after all user stories complete.

---

**Total Tasks**: 141 tasks organized across 8 phases
**MVP Tasks**: T001-T084 (84 tasks for User Story 1)
**Estimated Effort**: 
- MVP (User Story 1): 2-3 weeks for experienced developer
- Full Feature (All 5 user stories): 4-6 weeks
- With team parallelization: 3-4 weeks total
