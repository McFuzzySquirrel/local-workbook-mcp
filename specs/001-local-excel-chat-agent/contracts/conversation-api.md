# Conversation API Contract

**Feature**: Local Excel Conversational Agent  
**Date**: October 22, 2025

## Overview

This document defines the API contract between Blazor UI components and the ExcelAgentService. This is the primary interface for the chat functionality.

---

## IExcelAgentService

Main service interface for agent interactions.

### ProcessQueryAsync

Processes a user query and returns an agent response.

```csharp
Task<AgentResponse> ProcessQueryAsync(
    string query, 
    WorkbookSession session,
    CancellationToken cancellationToken = default
)
```

**Parameters**:
- `query`: User's natural language question
- `session`: Current workbook session with context
- `cancellationToken`: For timeout/cancellation support

**Returns**: `AgentResponse` with content, table data, or error

**Throws**:
- `ArgumentNullException`: If query or session is null
- `OperationCanceledException`: If cancelled or timeout exceeded
- `InvalidOperationException`: If no workbook loaded and query requires one

**Behavior**:
1. Validates query is not empty
2. Checks if workbook required based on query intent
3. Adds user turn to conversation history
4. Invokes Semantic Kernel with context window
5. SK calls appropriate plugins based on query
6. Formats response (text or table)
7. Adds assistant turn to history
8. Returns AgentResponse

**Error Handling**:
- LLM unresponsive → Return error response with retry option
- MCP tool failure → Return error with sanitized message
- Timeout (30s) → Cancel and return timeout error
- Ambiguous query → Return clarification request

---

### LoadWorkbookAsync

Loads a new workbook and updates session context.

```csharp
Task<WorkbookContext> LoadWorkbookAsync(
    string filePath,
    WorkbookSession session,
    CancellationToken cancellationToken = default
)
```

**Parameters**:
- `filePath`: Absolute path to .xlsx file
- `session`: Current session to update
- `cancellationToken`: For cancellation support

**Returns**: `WorkbookContext` with metadata or error info

**Throws**:
- `ArgumentNullException`: If filePath or session is null
- `FileNotFoundException`: If file doesn't exist
- `InvalidDataException`: If file is not valid .xlsx

**Behavior**:
1. Validates file path and existence
2. Checks file extension is .xlsx
3. Starts MCP server with workbook path
4. Retrieves workbook metadata from MCP
5. Creates WorkbookContext
6. Updates session.CurrentContext
7. Inserts system message marker in conversation
8. Returns WorkbookContext

**Error Handling**:
- File not found → Return context with IsValid=false, error message
- Corrupted file → Return context with IsValid=false, sanitized error
- Password-protected → Return context with error indicating not supported
- MCP server failure → Return context with error and suggested action

---

### ClearConversationAsync

Clears conversation history and starts fresh session.

```csharp
Task ClearConversationAsync(WorkbookSession session)
```

**Parameters**:
- `session`: Session to clear

**Returns**: Completed task

**Throws**: None (idempotent)

**Behavior**:
1. Clears session.ConversationHistory
2. Clears session.ContextWindow
3. Resets session.LastActivityAt to now
4. Does NOT unload workbook (CurrentContext remains)

---

### GetSuggestedQueriesAsync

Generates relevant follow-up questions based on current context.

```csharp
Task<List<string>> GetSuggestedQueriesAsync(
    WorkbookSession session,
    int maxSuggestions = 3
)
```

**Parameters**:
- `session`: Current session with workbook and conversation
- `maxSuggestions`: Number of suggestions to return

**Returns**: List of suggested query strings

**Behavior**:
1. Analyzes current workbook metadata
2. Reviews recent conversation turns
3. Generates contextually relevant suggestions
4. Returns list (may be empty if no good suggestions)

**Examples**:
- If workbook just loaded: ["What sheets are in this workbook?", "Show me a list of all tables"]
- After viewing Sales data: ["What's the total revenue?", "Show me top 10 sales", "Compare to last quarter"]

---

## IConversationManager

Manages conversation history and context window.

### AddUserTurn

Adds a user query to conversation.

```csharp
void AddUserTurn(string message, string correlationId)
```

**Parameters**:
- `message`: User's query text
- `correlationId`: Unique ID for tracking

**Behavior**:
1. Creates ConversationTurn with role="user"
2. Adds to full history
3. Adds to rolling context window (evicts oldest if > 20)
4. Updates session.LastActivityAt

---

### AddAssistantTurn

Adds an agent response to conversation.

```csharp
void AddAssistantTurn(string message, string correlationId, ContentType contentType)
```

**Parameters**:
- `message`: Agent's response text
- `correlationId`: Matches the query
- `contentType`: Type of response (Text, Table, Error, etc.)

**Behavior**: Same as AddUserTurn but with role="assistant"

---

### AddSystemMessage

Adds a system notification to conversation.

```csharp
void AddSystemMessage(string message)
```

**Parameters**:
- `message`: System message (e.g., "Workbook changed to Budget.xlsx")

**Behavior**: Adds turn with role="system", does NOT add to context window (UI display only)

---

### GetContextForLLM

Retrieves current context window for Semantic Kernel.

```csharp
ChatHistory GetContextForLLM()
```

**Returns**: SK ChatHistory with last 20 turns

**Behavior**: Returns the rolling window, not full history

---

### GetFullHistory

Retrieves complete conversation for UI display.

```csharp
List<ConversationTurn> GetFullHistory()
```

**Returns**: All turns including system messages

---

## IResponseFormatter

Formats agent responses for UI display.

### FormatAsHtmlTable

Converts table data to HTML table string.

```csharp
string FormatAsHtmlTable(TableData tableData)
```

**Parameters**:
- `tableData`: Structured table with columns and rows

**Returns**: HTML string with styled table

**HTML Structure**:
```html
<table class="excel-data-table">
  <thead>
    <tr><th>Column1</th><th>Column2</th>...</tr>
  </thead>
  <tbody>
    <tr><td>Value1</td><td>Value2</td>...</tr>
    ...
  </tbody>
</table>
```

**Behavior**:
- Escapes HTML in cell values for security
- Applies CSS classes for styling
- Truncates very large cells (>100 chars) with ellipsis
- Adds row count footer if table truncated

---

### FormatAsText

Formats a text response with markdown-like styling.

```csharp
string FormatAsText(string content)
```

**Parameters**:
- `content`: Plain text or simple markdown

**Returns**: Formatted HTML string

**Behavior**:
- Preserves line breaks as `<br>`
- Converts \*\*bold\*\* to `<strong>`
- Converts \*italic\* to `<em>`
- Converts `code` to `<code>`
- Converts links to `<a>` tags

---

### SanitizeErrorMessage

Sanitizes error for user display.

```csharp
SanitizedError SanitizeErrorMessage(Exception exception, string correlationId)
```

**Parameters**:
- `exception`: The actual exception
- `correlationId`: Tracking ID

**Returns**: `SanitizedError` with generic message

**Behavior**:
- Maps exception types to ErrorCode enum
- Generates user-friendly message
- Removes sensitive data (paths, sheet names, etc.)
- Adds suggested action based on error type
- Logs full exception details to file

---

## Component Events

### ChatComponentEvents

Events emitted by Blazor chat components.

#### OnQuerySubmitted

```csharp
EventCallback<string> OnQuerySubmitted { get; set; }
```

**Emitted**: When user submits a query via chat input

**Payload**: User's query string

---

#### OnWorkbookSelected

```csharp
EventCallback<string> OnWorkbookSelected { get; set; }
```

**Emitted**: When user selects a workbook file

**Payload**: Absolute file path

---

#### OnClearHistory

```csharp
EventCallback OnClearHistory { get; set; }
```

**Emitted**: When user clicks "Clear History" button

**Payload**: None

---

#### OnRetryQuery

```csharp
EventCallback OnRetryQuery { get; set; }
```

**Emitted**: When user clicks retry on an error

**Payload**: None (retries last query)

---

## Error Propagation

### Error Flow

```
Exception in service layer
    ↓
Catch in ExcelAgentService
    ↓
Log full details with correlationId
    ↓
Call IResponseFormatter.SanitizeErrorMessage
    ↓
Return AgentResponse with ContentType=Error
    ↓
Component displays SanitizedError
    ↓
User sees generic message + correlationId
    ↓
User can retry if CanRetry=true
```

### Example Error Handling in Component

```csharp
private async Task HandleQuerySubmit(string query)
{
    try
    {
        _isProcessing = true;
        StateHasChanged();
        
        var response = await _agentService.ProcessQueryAsync(
            query, 
            _session, 
            _cts.Token
        );
        
        if (response.ContentType == ContentType.Error)
        {
            // Display error with retry option
            _showRetryButton = response.Error.CanRetry;
        }
        
        _messages.Add(response);
    }
    catch (OperationCanceledException)
    {
        // User cancelled or timeout
        _messages.Add(new AgentResponse 
        { 
            ContentType = ContentType.Error,
            Error = new SanitizedError 
            { 
                Message = "Query was cancelled or timed out.",
                CanRetry = true
            }
        });
    }
    finally
    {
        _isProcessing = false;
        StateHasChanged();
    }
}
```

---

## Dependency Injection Setup

```csharp
// In Program.cs
builder.Services.AddScoped<IExcelAgentService, ExcelAgentService>();
builder.Services.AddScoped<IConversationManager, ConversationManager>();
builder.Services.AddSingleton<IResponseFormatter, ResponseFormatter>();

// Semantic Kernel
builder.Services.AddKernel()
    .AddOpenAIChatCompletion(/* ... */)
    .Plugins.AddFromObject<WorkbookStructurePlugin>()
    .Plugins.AddFromObject<WorkbookSearchPlugin>()
    .Plugins.AddFromObject<DataRetrievalPlugin>();

// Existing services
builder.Services.AddSingleton<IMcpClient, McpClientHost>();
builder.Services.AddHttpClient<ILlmStudioClient, LlmStudioClient>();

// Session state (per circuit)
builder.Services.AddScoped<WorkbookSession>();
```

---

## Testing Contracts

### Mocking for Unit Tests

```csharp
// Mock agent service for component tests
public class MockExcelAgentService : IExcelAgentService
{
    public Task<AgentResponse> ProcessQueryAsync(...)
    {
        return Task.FromResult(new AgentResponse
        {
            Content = "Mock response",
            ContentType = ContentType.Text,
            ProcessingTimeMs = 100
        });
    }
    
    // ... other methods
}
```

### Contract Tests

Each interface method should have tests verifying:
- Input validation (null checks, empty strings)
- Expected return types and structure
- Error case handling
- Timeout behavior
- Cancellation support
