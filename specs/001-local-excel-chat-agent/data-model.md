# Data Model

**Feature**: Local Excel Conversational Agent  
**Date**: October 22, 2025  
**Phase**: 1 - Design & Contracts

## Core Entities

### ConversationTurn

Represents a single message exchange in the conversation (user query or agent response).

**Fields**:
- `Id` (Guid): Unique identifier for the turn
- `Role` (string): "user", "assistant", or "system"
- `Content` (string): The message text or structured data
- `Timestamp` (DateTimeOffset): When the turn occurred
- `CorrelationId` (string): Links related operations across logs
- `ContentType` (ContentType enum): Text, Table, Error, SystemMessage
- `Metadata` (Dictionary<string, object>?): Optional additional data

**Validation Rules**:
- Role must be one of: "user", "assistant", "system"
- Content cannot be null or empty for user/assistant roles
- Timestamp must be in UTC
- CorrelationId must be valid GUID format

**State Transitions**:
- Created → Stored in conversation history → Archived (when evicted from rolling window)

**Relationships**:
- Part of WorkbookSession
- May reference WorkbookContext for system messages

---

### WorkbookContext

Represents the currently loaded workbook and its metadata.

**Fields**:
- `WorkbookPath` (string): Absolute file path to the .xlsx file
- `WorkbookName` (string): Display name (filename without path)
- `LoadedAt` (DateTimeOffset): When the workbook was loaded
- `Metadata` (WorkbookMetadata): Structured info from MCP server
  - `SheetNames` (List<string>): All worksheet names
  - `TableNames` (Dictionary<string, List<string>>): Tables per sheet
  - `RowCounts` (Dictionary<string, int>): Row count per sheet
  - `ColumnCounts` (Dictionary<string, int>): Column count per sheet
- `IsValid` (bool): Whether workbook loaded successfully
- `ErrorMessage` (string?): Sanitized error if load failed

**Validation Rules**:
- WorkbookPath must be absolute path
- WorkbookPath must point to .xlsx file
- WorkbookName must not contain path separators
- If IsValid is false, ErrorMessage must be set

**State Transitions**:
- Unloaded → Loading → Loaded (IsValid=true) → Replaced (new workbook loaded)
- Unloaded → Loading → Failed (IsValid=false)

**Relationships**:
- One per WorkbookSession at a time
- History of previous contexts maintained for multi-workbook sessions

---

### WorkbookSession

Represents an active user session with conversation history and workbook state.

**Fields**:
- `SessionId` (Guid): Unique session identifier
- `StartedAt` (DateTimeOffset): Session creation time
- `LastActivityAt` (DateTimeOffset): Most recent interaction time
- `CurrentContext` (WorkbookContext?): Active workbook, null if none loaded
- `ConversationHistory` (List<ConversationTurn>): Full conversation (all turns)
- `ContextWindow` (ChatHistory): SK ChatHistory with last 20 turns for LLM
- `PreviousContexts` (List<WorkbookContext>): History of loaded workbooks

**Validation Rules**:
- SessionId must be unique
- ConversationHistory cannot be null (can be empty)
- ContextWindow.Count <= 20 turns
- LastActivityAt >= StartedAt

**State Transitions**:
- Created → Active (user interacting) → Idle (no activity) → Expired (timeout)
- Session persists in-memory until browser disconnect or explicit clear

**Relationships**:
- Contains multiple ConversationTurns
- References current and previous WorkbookContexts
- Owned by single Blazor circuit/user connection

---

### AgentResponse

Represents the structured response from the Semantic Kernel agent.

**Fields**:
- `ResponseId` (Guid): Unique response identifier
- `CorrelationId` (string): Links to the query that triggered this response
- `Content` (string): Primary response text
- `ContentType` (ContentType enum): Text, Table, Error, Clarification
- `TableData` (TableData?): Structured table if ContentType is Table
  - `Columns` (List<string>): Column headers
  - `Rows` (List<List<string>>): Data rows
  - `Metadata` (TableMetadata): Sheet name, row count, is truncated
- `Suggestions` (List<string>): Follow-up query suggestions
- `ToolsInvoked` (List<ToolInvocation>): MCP tools called by agent
- `ProcessingTimeMs` (int): Time taken to generate response
- `ModelUsed` (string): LLM model identifier
- `Error` (SanitizedError?): Error info if response is error

**Validation Rules**:
- If ContentType is Table, TableData must not be null
- If ContentType is Error, Error must not be null
- ProcessingTimeMs must be >= 0
- ToolsInvoked list accurately reflects MCP calls made

**State Transitions**:
- Processing → Completed → Displayed → Archived

**Relationships**:
- Created in response to ConversationTurn (user query)
- Becomes next ConversationTurn (assistant role) in history

---

### ToolInvocation

Represents a call to an MCP tool via Semantic Kernel plugin.

**Fields**:
- `ToolName` (string): MCP tool identifier (e.g., "excel-list-structure")
- `PluginMethod` (string): SK plugin method name
- `InputParameters` (Dictionary<string, object>): Parameters passed to tool
- `OutputSummary` (string): Brief summary of result (not full data)
- `InvokedAt` (DateTimeOffset): Execution timestamp
- `DurationMs` (int): Execution time
- `Success` (bool): Whether tool call succeeded
- `ErrorMessage` (string?): Sanitized error if failed

**Validation Rules**:
- ToolName must match known MCP tools
- If Success is false, ErrorMessage must be set
- DurationMs must be >= 0

**State Transitions**:
- Queued → Executing → Completed (Success=true) → Logged
- Queued → Executing → Failed (Success=false) → Logged

**Relationships**:
- Part of AgentResponse.ToolsInvoked collection
- Logged for observability and troubleshooting

---

### SanitizedError

Represents a user-safe error message with troubleshooting reference.

**Fields**:
- `Message` (string): Generic user-friendly error message
- `CorrelationId` (string): Links to detailed log entry
- `Timestamp` (DateTimeOffset): When error occurred
- `ErrorCode` (ErrorCode enum): Categorized error type
  - `WorkbookLoadFailed`
  - `QueryTimeout`
  - `ModelUnresponsive`
  - `InvalidQuery`
  - `McpToolError`
  - `UnknownError`
- `CanRetry` (bool): Whether user can retry the operation
- `SuggestedAction` (string?): Helpful guidance for user

**Validation Rules**:
- Message must not contain sensitive data (file paths, sheet names, cell values)
- CorrelationId must be valid GUID format
- If CanRetry is true, SuggestedAction should guide retry

**Relationships**:
- Part of AgentResponse when ContentType is Error
- CorrelationId links to detailed AgentLogger entries

---

## Enumerations

### ContentType

```csharp
public enum ContentType
{
    Text,           // Plain text response
    Table,          // Structured data table
    Error,          // Error message
    SystemMessage,  // System notification (e.g., workbook switch marker)
    Clarification   // Agent asking for user clarification
}
```

### ErrorCode

```csharp
public enum ErrorCode
{
    WorkbookLoadFailed,      // Could not open/read workbook
    QueryTimeout,            // Query exceeded 30s limit
    ModelUnresponsive,       // LLM server not responding
    InvalidQuery,            // Query could not be parsed/understood
    McpToolError,            // MCP tool invocation failed
    UnknownError             // Unexpected error
}
```

---

## Data Flow

### User Query Flow

```
User types query in Blazor UI
    ↓
ChatMessage component emits event
    ↓
Chat.razor.cs receives query
    ↓
ExcelAgentService.ProcessQueryAsync(query, sessionState)
    ↓
ConversationManager adds user turn to history + context window
    ↓
Semantic Kernel processes query with context
    ↓
SK invokes plugin methods (WorkbookStructurePlugin, etc.)
    ↓
Plugin calls McpClientHost → MCP Server → Returns data
    ↓
SK generates response using LLM + tool results
    ↓
ResponseFormatter converts to AgentResponse
    ↓
ConversationManager adds assistant turn
    ↓
AgentResponse returned to Chat.razor.cs
    ↓
ChatMessage component renders response
```

### Workbook Load Flow

```
User selects file in WorkbookSelector component
    ↓
Chat.razor.cs receives file path
    ↓
ExcelAgentService.LoadWorkbookAsync(path)
    ↓
McpClientHost starts MCP server with workbook
    ↓
MCP server returns WorkbookMetadata
    ↓
WorkbookContext created with metadata
    ↓
WorkbookSessionState.LoadNewWorkbook() inserts marker
    ↓
CurrentContext updated to new workbook
    ↓
UI displays confirmation + workbook summary
```

---

## Persistence & Lifecycle

### In-Memory Storage (Session Scope)

**Stored**:
- `WorkbookSession` instance per Blazor circuit
- Full `ConversationHistory` for UI display
- Rolling 20-turn `ContextWindow` for LLM
- Current `WorkbookContext`

**Lifecycle**:
- Created: When user first accesses chat page
- Active: During user interaction
- Cleared: When user clicks "Clear History" or circuit disconnects

### File-Based Storage (Local Logs)

**Stored**:
- Structured JSON Lines logs in `logs/agent-{date}.log`
- Each line: { timestamp, correlationId, event, details }
- Events: AgentQuery, ToolInvoked, ResponseGenerated, Error

**Lifecycle**:
- Appended: Real-time as events occur
- Retained: Per application settings (default: 30 days)
- Rotated: Daily log files

---

## Validation & Constraints

### Rolling Window Management

```csharp
// Maximum context turns for LLM
public const int MaxContextTurns = 20;

// Eviction strategy: FIFO (oldest first)
while (ContextWindow.Count > MaxContextTurns)
{
    ContextWindow.RemoveAt(0);
}
```

### Data Size Limits

- **TableData.Rows**: Max 1000 rows per response (pagination for larger)
- **Content**: Max 10,000 characters (truncate with "..." indicator)
- **ConversationHistory**: No limit (in-memory, session-scoped)
- **Logs**: 100MB daily files, 30 days retention

### Sanitization Rules

**Never include in UI messages**:
- Absolute file paths → Use filename only
- Sheet/table names in error context → Use "Sheet not found" not "Sheet 'Confidential' not found"
- Cell values in errors → Use "Invalid data" not "Cell A1 contains 'Secret'"
- Stack traces → Log only, never display

---

## Technology Mapping

### Semantic Kernel Integration

```csharp
// SK ChatHistory stores ContextWindow
using Microsoft.SemanticKernel.ChatCompletion;

public class ConversationManager
{
    private ChatHistory _contextWindow = new();
    
    public void AddUserMessage(string message)
    {
        _contextWindow.AddUserMessage(message);
        EnforceWindowLimit();
    }
    
    public void AddAssistantMessage(string message)
    {
        _contextWindow.AddAssistantMessage(message);
        EnforceWindowLimit();
    }
    
    public ChatHistory GetContextForLLM() => _contextWindow;
}
```

### MCP Data Mapping

```csharp
// MCP WorkbookMetadata → WorkbookContext
public WorkbookContext FromMcpMetadata(string path, WorkbookMetadata mcp)
{
    return new WorkbookContext
    {
        WorkbookPath = path,
        WorkbookName = Path.GetFileName(path),
        LoadedAt = DateTimeOffset.UtcNow,
        Metadata = mcp,
        IsValid = true
    };
}
```

---

## Next Steps

Proceed to contracts definition:
- SK plugin method signatures
- Blazor component event interfaces
- API between Chat.razor.cs and ExcelAgentService
