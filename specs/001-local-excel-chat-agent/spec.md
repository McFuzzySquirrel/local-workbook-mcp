# Feature Specification: Local Excel Conversational Agent

**Feature Branch**: `001-local-excel-chat-agent`  
**Created**: October 22, 2025  
**Status**: Draft  
**Input**: User description: "build an agent that will allow a user to have a conversation with their excel workbook using a local small language model that has tool calling capability, the user should be able to query the workbook and get insights. The user interface should be simple and intuitive and run on a local device"

## Clarifications

### Session 2025-10-22

- Q: When the local language model becomes unresponsive or returns an error, how should the system behave? → A: Show error message, retain workbook and conversation history, allow retry of the last query
- Q: When the agent displays tabular data from the workbook (e.g., "Show me the first 10 rows"), how should the data be formatted in the chat interface? → A: Formatted HTML table with headers, borders, and proper column alignment
- Q: When displaying error messages to users, how should the system handle potentially sensitive information from workbook content or file paths? → A: Show sanitized error messages (e.g., "Sheet not found" without revealing sheet name), log full details locally for troubleshooting
- Q: When a user loads a different workbook during an active session (FR-015), what should happen to the conversation history? → A: Keep conversation history, insert a clear marker showing workbook switch, agent knows only new workbook is active
- Q: How many conversation turns should the system maintain in context when processing new queries? → A: Keep most recent 20 turns (10 user queries + 10 agent responses) in rolling window

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Basic Workbook Querying (Priority: P1)

A user opens the chat interface, loads an Excel workbook from their local device, and asks natural language questions about the data. The agent responds with insights extracted from the workbook, such as sheet names, table summaries, or specific data values.

**Why this priority**: This is the core functionality that delivers immediate value. Users can interact with their spreadsheets conversationally without needing to manually navigate complex workbooks or write queries.

**Independent Test**: Can be fully tested by loading a sample workbook, asking "What sheets are in this workbook?" and verifying the agent returns accurate sheet names. Delivers standalone value by making workbook content discoverable through conversation.

**Acceptance Scenarios**:

1. **Given** a user has opened the chat interface, **When** they select an Excel workbook from their device, **Then** the system confirms the workbook is loaded and ready for queries
2. **Given** a workbook is loaded, **When** the user asks "What sheets exist in this workbook?", **Then** the agent lists all worksheet names
3. **Given** a workbook is loaded, **When** the user asks "Show me the first 10 rows of the Sales table", **Then** the agent displays the requested data in a readable format
4. **Given** a workbook is loaded, **When** the user asks "What is the total revenue in the Q4 worksheet?", **Then** the agent calculates and returns the sum
5. **Given** no workbook is loaded, **When** the user asks a question, **Then** the agent prompts them to load a workbook first

---

### User Story 2 - Multi-Turn Conversations with Context (Priority: P2)

A user engages in a multi-turn conversation where follow-up questions reference previous queries. The agent maintains conversation context and understands references like "it", "that table", or "those values".

**Why this priority**: Enables natural, flowing conversations rather than isolated single queries. This significantly improves user experience by reducing repetition and making interactions feel more intuitive.

**Independent Test**: Can be tested by loading a workbook, asking "What's in the Sales sheet?", then following up with "Show me the top 5 rows" without re-specifying the sheet name. Delivers value by enabling natural conversation flow.

**Acceptance Scenarios**:

1. **Given** the user asked "What tables are in the Budget worksheet?", **When** they follow up with "Show me the first one", **Then** the agent displays data from the first table mentioned in the previous response
2. **Given** the user asked about revenue data, **When** they ask "What about last quarter?", **Then** the agent understands the context and provides comparative data
3. **Given** a conversation has multiple turns, **When** the user asks "Can you summarize what we've discussed?", **Then** the agent provides a recap of queries and findings
4. **Given** the user starts a new topic, **When** they say "Let's look at something else", **Then** the agent acknowledges the context switch and is ready for new queries

---

### User Story 3 - Cross-Sheet Data Insights (Priority: P2)

A user asks questions that require analyzing data across multiple worksheets or tables. The agent correlates information from different sources within the workbook to provide comprehensive insights.

**Why this priority**: Demonstrates the agent's analytical capabilities beyond simple data retrieval. Users often need to understand relationships between different parts of their workbook.

**Independent Test**: Can be tested by loading a workbook with multiple related sheets (e.g., Sales and Returns), asking "Are there any products that appear in both Sales and Returns sheets?", and verifying the agent identifies common entries. Delivers value by automating cross-reference analysis.

**Acceptance Scenarios**:

1. **Given** a workbook has Sales and Inventory sheets, **When** the user asks "Which products have low inventory but high sales?", **Then** the agent analyzes both sheets and identifies matching items
2. **Given** multiple worksheets contain date-based data, **When** the user asks "Compare revenue trends across all quarterly sheets", **Then** the agent aggregates and presents comparative insights
3. **Given** the user asks "Find all entries mentioning Project X across the entire workbook", **When** the search spans multiple sheets, **Then** the agent returns results from all relevant locations

---

### User Story 4 - Simple Data Filtering and Searching (Priority: P3)

A user requests specific subsets of data based on criteria like value ranges, text matching, or date filters. The agent retrieves and displays only the relevant rows.

**Why this priority**: Extends querying capabilities to support more targeted analysis. While not critical for MVP, this significantly enhances usefulness for users working with large datasets.

**Independent Test**: Can be tested by loading a sales workbook and asking "Show me all sales over $10,000 in January", then verifying only matching rows are returned. Delivers value by eliminating manual filtering work.

**Acceptance Scenarios**:

1. **Given** a workbook contains transaction data, **When** the user asks "Show sales between $5,000 and $15,000", **Then** the agent displays only rows matching that range
2. **Given** a workbook has customer data, **When** the user asks "Find all entries for customers in California", **Then** the agent filters and returns matching records
3. **Given** the user specifies multiple criteria, **When** they ask "Show Q3 sales over $20,000 from the West region", **Then** the agent applies all filters correctly

---

### User Story 5 - Exporting Conversation Insights (Priority: P4)

A user wants to save or share insights gained from their conversation with the workbook. The agent provides options to export summaries, specific data views, or the conversation history.

**Why this priority**: Nice-to-have feature that supports collaboration and documentation. Not essential for core functionality but improves practical usability.

**Independent Test**: Can be tested by conducting a query session, then requesting "Export our conversation", and verifying a readable summary file is created. Delivers value by preserving insights for later reference or sharing.

**Acceptance Scenarios**:

1. **Given** a user has completed several queries, **When** they ask "Export this conversation", **Then** the agent generates a readable summary file
2. **Given** the agent displayed a data table, **When** the user asks "Save that table", **Then** the agent exports the specific data view
3. **Given** the user wants to document findings, **When** they request "Create a summary of our key insights", **Then** the agent generates a concise report of important findings

---

### Edge Cases

- What happens when the selected Excel file is corrupted or cannot be read? System displays error message and prompts user to select a different file.
- How does the system handle queries about non-existent sheets or tables? Agent responds indicating the sheet/table was not found and lists available options.
- What happens if the workbook is password-protected? System detects protection during load and prompts user that password-protected files are not supported.
- How does the agent respond to ambiguous queries that could reference multiple sheets or tables? Agent asks clarifying question presenting the matching options.
- What happens when the workbook is extremely large (thousands of rows/columns)? System may take longer to process (within 30-second limit per SC-006) or suggest narrowing the query scope.
- How does the system handle queries while the workbook file is being modified by another application? System works with the version loaded into memory; changes require reloading the workbook.
- What happens if the local language model becomes unresponsive or unavailable? System displays error message, retains workbook and conversation history, and allows user to retry the query.
- How does the agent handle queries that would require excessive computation time? System enforces timeout (30 seconds per SC-006), then prompts user to refine query.
- What happens when the user asks questions outside the scope of workbook analysis (e.g., general knowledge questions)? Agent politely declines and reminds user it can only answer questions about the loaded workbook.

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: System MUST allow users to select and load Excel workbook files (.xlsx) from their local device
- **FR-002**: System MUST display a simple chat interface where users can type natural language queries
- **FR-003**: System MUST process user queries using a local small language model with tool calling capabilities
- **FR-004**: System MUST expose MCP tools for workbook analysis that the language model can invoke (list structure, search, preview tables, retrieve data)
- **FR-005**: System MUST display agent responses in a conversational, readable format within the chat interface
- **FR-005a**: When displaying tabular data from the workbook, system MUST render data as formatted HTML tables with headers, borders, and proper column alignment
- **FR-006**: System MUST maintain conversation history for the duration of the session
- **FR-006a**: System MUST maintain the most recent 20 conversation turns (10 user queries and 10 agent responses) in a rolling window for context when processing new queries; older turns may be archived for display purposes but not included in language model context
- **FR-007**: System MUST support queries about workbook structure (sheet names, table names, column headers)
- **FR-008**: System MUST support queries about specific data values (cell contents, row data, calculated aggregations)
- **FR-009**: System MUST support search operations across workbook content
- **FR-010**: System MUST handle errors gracefully and provide user-friendly error messages when queries cannot be completed
- **FR-010a**: Error messages displayed to users MUST be sanitized to avoid exposing sensitive workbook content, sheet names, or file paths; full error details MUST be logged locally for troubleshooting
- **FR-011**: System MUST run entirely on the user's local device without requiring internet connectivity for core functionality
- **FR-012**: System MUST preserve user privacy by never transmitting workbook data outside the local environment
- **FR-013**: System MUST validate that loaded files are valid Excel workbooks before processing
- **FR-014**: Users MUST be able to clear conversation history and start a new session
- **FR-015**: Users MUST be able to load a different workbook during an active session
- **FR-015a**: When a user loads a different workbook, system MUST retain conversation history, insert a clear visual marker indicating the workbook switch, and ensure the agent processes subsequent queries against only the newly loaded workbook
- **FR-016**: System MUST indicate when it is processing a query (loading state)
- **FR-017**: System MUST provide clear feedback when a workbook is successfully loaded
- **FR-018**: System MUST handle queries that reference context from previous conversation turns
- **FR-019**: System MUST limit response lengths to prevent overwhelming the user interface
- **FR-020**: System MUST identify when queries are ambiguous and ask clarifying questions
- **FR-021**: When the language model becomes unresponsive or returns an error, system MUST display an error message, retain the loaded workbook and conversation history, and provide the user an option to retry the last query

### Key Entities

- **Workbook Session**: Represents an active user session with a loaded Excel workbook, including conversation history and current context; when a new workbook is loaded, the session continues with history preserved and a marker inserted to indicate the workbook change
- **Query**: A natural language question or request from the user about workbook data
- **Response**: The agent's answer to a query, which may include text explanations, data tables (rendered as HTML tables with headers and borders), or clarifying questions
- **Workbook Metadata**: Information about the loaded workbook including sheet names, table names, row/column counts, and data ranges
- **Tool Invocation**: A specific MCP tool call made by the language model to retrieve workbook information (e.g., list structure, search, preview)

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001**: Users can load an Excel workbook and receive confirmation within 5 seconds for workbooks up to 50MB
- **SC-002**: Users can ask basic queries (sheet names, table lists) and receive accurate responses within 10 seconds
- **SC-003**: The agent correctly answers at least 90% of queries about workbook structure (sheets, tables, columns)
- **SC-004**: The agent successfully retrieves specific data values with 95% accuracy when the query is unambiguous
- **SC-005**: Users can complete a typical analysis workflow (load workbook, ask 3-5 questions, get insights) in under 3 minutes
- **SC-006**: The system handles workbooks with up to 100 sheets and 100,000 total rows without crashing or excessive delays (queries complete within 30 seconds)
- **SC-007**: 80% of users successfully complete their first query without requiring help documentation
- **SC-008**: The conversation interface maintains context across 5+ consecutive related queries without requiring users to repeat information (system maintains rolling window of most recent 20 turns)
- **SC-009**: The system provides helpful error messages for 100% of common failure scenarios (file not found, corrupted workbook, invalid query)
- **SC-010**: The application runs successfully on typical consumer devices with 8GB RAM and standard processors without requiring GPU acceleration

## Assumptions

- The existing ExcelMcp.Server and ExcelMcp.Contracts infrastructure can be leveraged for MCP tool implementation
- Users will have LM Studio or a compatible OpenAI-like local inference server installed with a suitable small language model (e.g., Phi-4, Mistral 7B, or similar models under 10GB)
- The local language model supports function/tool calling capabilities (or can be prompted to generate structured tool invocation requests)
- Excel workbooks are stored locally in .xlsx format (not .xls or other legacy formats)
- Users have basic familiarity with their workbook structure (know general content, not necessarily specific details)
- The chat interface will be web-based and accessed through a local browser (localhost)
- Conversation history is stored in memory and cleared when the session ends (no persistent storage required for MVP)
- The agent will focus on read-only operations; writing back to workbooks is out of scope for this feature
- Standard web browser security allows file selection from local filesystem
- Users understand that the quality of responses depends on the capabilities of their chosen language model

## Dependencies

- Existing ExcelMcp.Server component for workbook reading and MCP tool exposure
- Local language model inference server (e.g., LM Studio, Ollama, or similar)
- ClosedXML library for Excel file parsing
- .NET 9.0+ runtime for server components

## Constraints

- Must operate entirely offline after initial setup (no cloud API calls for workbook analysis)
- Must work with small language models (under 10GB) that can run on consumer hardware
- Must not require specialized AI hardware (GPUs optional but not required)
- Must respect user privacy - no telemetry or data transmission
- Must work with standard .xlsx files (no macro-enabled workbooks initially)
