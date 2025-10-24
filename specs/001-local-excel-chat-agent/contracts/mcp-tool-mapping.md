# MCP Tool to SK Plugin Mapping

**Feature**: Local Excel Conversational Agent  
**Date**: October 22, 2025

## Overview

This document maps existing ExcelMcp.Server MCP tools to Semantic Kernel plugin methods, showing how the agent bridges MCP protocol with SK orchestration.

---

## Existing MCP Tools

### excel-list-structure

**Description**: Returns metadata about the workbook structure

**Input Parameters**: None (uses loaded workbook)

**Output Schema**:
```json
{
  "sheets": [
    {
      "name": "string",
      "rowCount": number,
      "columnCount": number,
      "tables": ["string"]
    }
  ]
}
```

**SK Plugin Methods Using This Tool**:
- `WorkbookStructurePlugin.ListWorkbookStructure()`
- `WorkbookStructurePlugin.GetSheetNames()`
- `WorkbookStructurePlugin.GetTableInfo(sheetName)`

---

### excel-search

**Description**: Searches for text across all worksheets

**Input Parameters**:
```json
{
  "query": "string",
  "maxResults": number (optional, default 50)
}
```

**Output Schema**:
```json
{
  "matches": [
    {
      "sheetName": "string",
      "cellAddress": "string",
      "value": "string",
      "row": number,
      "column": number
    }
  ],
  "totalMatches": number,
  "truncated": boolean
}
```

**SK Plugin Methods Using This Tool**:
- `WorkbookSearchPlugin.SearchWorkbook(searchText, maxResults)`
- `WorkbookSearchPlugin.SearchInSheet(sheetName, searchText, maxResults)`

---

### excel-preview-table

**Description**: Retrieves rows from a table or worksheet

**Input Parameters**:
```json
{
  "resourceUri": "excel://path/to/file.xlsx/SheetName/TableName",
  "rowCount": number (optional, default 10),
  "startRow": number (optional, default 0)
}
```

**Output Schema**:
```json
{
  "columns": ["string"],
  "rows": [["string"]],
  "totalRows": number,
  "startRow": number,
  "hasMore": boolean
}
```

**SK Plugin Methods Using This Tool**:
- `DataRetrievalPlugin.PreviewTable(name, rowCount, startRow)`
- `DataRetrievalPlugin.GetRowsInRange(sheetName, cellRange)`
- `DataRetrievalPlugin.CalculateAggregation(name, column, aggregationType)` (retrieves data, calculates client-side)

---

## MCP Resource URIs

### Format

`excel://[workbook-path]/[sheet-name]/[table-name]`

**Examples**:
- `excel://C:/Data/Sales.xlsx/Sheet1` - References a worksheet
- `excel://C:/Data/Sales.xlsx/Sheet1/SalesData` - References a named table

### Usage in SK Plugins

SK plugins construct resource URIs based on:
- Current workbook path from WorkbookContext
- Sheet/table names from user queries
- Pass URIs to MCP tools via McpClientHost

```csharp
public async Task<string> PreviewTable(string name, int rowCount)
{
    // Get current workbook path from session
    var workbookPath = _session.CurrentContext.WorkbookPath;
    
    // Construct resource URI
    var resourceUri = $"excel://{workbookPath}/{name}";
    
    // Call MCP tool
    var result = await _mcpClient.CallToolAsync(
        "excel-preview-table",
        new { resourceUri, rowCount }
    );
    
    return JsonSerializer.Serialize(result);
}
```

---

## MCP Client Adapter

### IMcpClient Interface

```csharp
public interface IMcpClient
{
    Task<JsonElement> CallToolAsync(
        string toolName, 
        object? parameters = null,
        CancellationToken cancellationToken = default
    );
    
    Task<JsonElement> GetResourceAsync(
        string resourceUri,
        CancellationToken cancellationToken = default
    );
    
    Task<bool> IsServerReadyAsync();
}
```

### McpClientHost Implementation

Existing `McpClientHost` class wraps stdio JSON-RPC communication with MCP server.

**Enhancements Needed**:
- Add timeout support (30s default)
- Add correlation ID to all requests
- Structured error responses on MCP failures
- Connection health monitoring

---

## Plugin to MCP Call Flow

### Example: User asks "What sheets are in the workbook?"

**Step 1**: Blazor Chat.razor receives query
```csharp
await _agentService.ProcessQueryAsync("What sheets are in the workbook?", _session);
```

**Step 2**: ExcelAgentService invokes Semantic Kernel
```csharp
var chatHistory = _conversationManager.GetContextForLLM();
chatHistory.AddUserMessage(query);

var result = await _kernel.InvokePromptAsync(
    "Answer the user's question about the Excel workbook using available tools.",
    new KernelArguments { ["chat_history"] = chatHistory }
);
```

**Step 3**: SK determines tool to call (via LLM function calling)
```json
{
  "tool_call": {
    "name": "list_workbook_structure",
    "parameters": {}
  }
}
```

**Step 4**: SK invokes WorkbookStructurePlugin.ListWorkbookStructure()
```csharp
public async Task<string> ListWorkbookStructure()
{
    // Call MCP tool
    var mcpResult = await _mcpClient.CallToolAsync("excel-list-structure");
    
    // Format response
    return JsonSerializer.Serialize(mcpResult);
}
```

**Step 5**: Plugin calls McpClientHost → MCP Server
```csharp
// McpClientHost sends JSON-RPC request
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "excel-list-structure",
    "arguments": {}
  }
}
```

**Step 6**: MCP Server processes and returns data
```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "result": {
    "content": [
      {
        "type": "text",
        "text": "{\"sheets\":[{\"name\":\"Sales\",\"rowCount\":1523,...}]}"
      }
    ]
  }
}
```

**Step 7**: Plugin returns data to SK
```csharp
return parsedResult; // JSON string
```

**Step 8**: SK generates natural language response
```
The workbook contains 2 sheets: Sales (1,523 rows) and Inventory (845 rows).
```

**Step 9**: ExcelAgentService formats and returns AgentResponse
```csharp
return new AgentResponse
{
    Content = skResult.ToString(),
    ContentType = ContentType.Text,
    ToolsInvoked = [new ToolInvocation { ToolName = "excel-list-structure", ... }],
    ProcessingTimeMs = elapsed
};
```

**Step 10**: Blazor component displays response

---

## Error Mapping

### MCP Error to SK Plugin Error

| MCP Error | SK Plugin Error Code | User Message |
|-----------|---------------------|--------------|
| Resource not found | SHEET_NOT_FOUND | "The requested sheet could not be found" |
| Invalid arguments | INVALID_INPUT | "The query parameters were invalid" |
| Timeout | TIMEOUT | "The operation took too long" |
| Server error | MCP_ERROR | "An error occurred accessing the workbook" |
| Server not ready | NO_WORKBOOK | "No workbook is currently loaded" |

### Implementation

```csharp
private ErrorResponse HandleMcpError(Exception ex, string correlationId)
{
    var errorCode = ex switch
    {
        McpResourceNotFoundException => "SHEET_NOT_FOUND",
        McpInvalidArgumentsException => "INVALID_INPUT",
        TimeoutException => "TIMEOUT",
        McpServerNotReadyException => "NO_WORKBOOK",
        _ => "MCP_ERROR"
    };
    
    return new ErrorResponse
    {
        Error = true,
        ErrorCode = errorCode,
        Message = GetUserFriendlyMessage(errorCode),
        CorrelationId = correlationId,
        CanRetry = IsRetryable(errorCode)
    };
}
```

---

## Tool Call Logging

Every MCP tool invocation should be logged for observability:

```csharp
_logger.LogInformation(
    "MCP tool invoked. Tool: {ToolName}, CorrelationId: {CorrelationId}, Duration: {DurationMs}ms",
    "excel-list-structure",
    correlationId,
    durationMs
);
```

**Log Entry Example** (JSON Lines):
```json
{
  "timestamp": "2025-10-22T14:30:00.123Z",
  "level": "Information",
  "correlationId": "abc-123",
  "event": "ToolInvoked",
  "tool": "excel-list-structure",
  "plugin": "WorkbookStructurePlugin",
  "method": "ListWorkbookStructure",
  "durationMs": 45,
  "success": true
}
```

---

## Performance Considerations

### Caching Strategy

Cache workbook metadata to avoid repeated MCP calls:

```csharp
public class WorkbookStructurePlugin
{
    private CachedStructure? _cache;
    
    public async Task<string> ListWorkbookStructure()
    {
        // Check cache validity
        if (_cache?.IsValid(_session.CurrentContext.LoadedAt) == true)
        {
            return _cache.Data;
        }
        
        // Fetch from MCP
        var result = await _mcpClient.CallToolAsync("excel-list-structure");
        
        // Update cache
        _cache = new CachedStructure(result, DateTime.UtcNow);
        
        return result;
    }
}
```

### Timeout Enforcement

All MCP calls must respect 30-second timeout:

```csharp
using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));

var result = await _mcpClient.CallToolAsync(
    toolName, 
    parameters, 
    cts.Token
);
```

---

## Testing MCP Integration

### Unit Tests with Mock MCP Client

```csharp
public class WorkbookStructurePluginTests
{
    [Fact]
    public async Task ListWorkbookStructure_CallsCorrectMcpTool()
    {
        // Arrange
        var mockMcp = new Mock<IMcpClient>();
        mockMcp.Setup(m => m.CallToolAsync("excel-list-structure", null, default))
               .ReturnsAsync(JsonDocument.Parse("{\"sheets\":[]}").RootElement);
        
        var plugin = new WorkbookStructurePlugin(mockMcp.Object);
        
        // Act
        await plugin.ListWorkbookStructure();
        
        // Assert
        mockMcp.Verify(m => m.CallToolAsync("excel-list-structure", null, default), Times.Once);
    }
}
```

### Integration Tests with Real MCP Server

```csharp
[Fact]
public async Task EndToEnd_LoadWorkbookAndQuery()
{
    // Arrange
    var mcpClient = new McpClientHost(/* test workbook path */);
    await mcpClient.StartAsync();
    
    var plugin = new WorkbookStructurePlugin(mcpClient);
    
    // Act
    var result = await plugin.ListWorkbookStructure();
    
    // Assert
    var data = JsonSerializer.Deserialize<WorkbookStructure>(result);
    Assert.NotNull(data);
    Assert.NotEmpty(data.Sheets);
}
```
