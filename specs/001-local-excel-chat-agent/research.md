# Research & Technology Decisions

**Feature**: Local Excel Conversational Agent  
**Date**: October 22, 2025  
**Phase**: 0 - Research & Technology Selection

## Research Questions Resolved

### Q1: How to integrate Semantic Kernel with existing MCP server?

**Decision**: Wrap MCP tools as Semantic Kernel plugins with JSON-RPC client

**Rationale**:
- Semantic Kernel's plugin architecture naturally maps to MCP tools
- Each MCP tool (excel-list-structure, excel-search, excel-preview-table) becomes a SK plugin method
- SK handles orchestration, context management, and LLM interaction
- MCP server remains unchanged, maintaining separation of concerns
- Existing McpClientHost can be adapted to work with SK plugins

**Implementation Pattern**:
```csharp
// SK Plugin wraps MCP client
public class WorkbookStructurePlugin
{
    private readonly IMcpClient _mcpClient;
    
    [KernelFunction("list_workbook_structure")]
    [Description("Lists all sheets and tables in the loaded workbook")]
    public async Task<string> ListStructure()
    {
        var result = await _mcpClient.CallToolAsync("excel-list-structure");
        return JsonSerializer.Serialize(result);
    }
}
```

**Alternatives Considered**:
- Direct LLM integration without SK: Rejected - would require custom orchestration logic and lose SK's planning capabilities
- Rebuilding MCP tools as native SK plugins: Rejected - duplicates existing working code and breaks MCP compatibility

---

### Q2: How to implement 20-turn rolling conversation window?

**Decision**: In-memory circular buffer with Semantic Kernel ChatHistory

**Rationale**:
- Semantic Kernel provides `ChatHistory` class for conversation management
- Implement custom `ConversationManager` with fixed-size rolling window
- Keep last 20 turns (10 user + 10 assistant) in SK ChatHistory for LLM context
- Store full history separately for UI display purposes
- Efficient memory usage - old turns automatically evicted
- No persistence required per spec (session-based only)

**Implementation Pattern**:
```csharp
public class ConversationManager
{
    private readonly ChatHistory _contextWindow = new();
    private readonly List<ConversationTurn> _fullHistory = new();
    private const int MaxContextTurns = 20;
    
    public void AddTurn(string role, string message)
    {
        // Add to full history for UI
        _fullHistory.Add(new ConversationTurn(role, message));
        
        // Add to rolling window
        if (role == "user")
            _contextWindow.AddUserMessage(message);
        else
            _contextWindow.AddAssistantMessage(message);
            
        // Evict oldest if exceeds limit
        while (_contextWindow.Count > MaxContextTurns)
            _contextWindow.RemoveAt(0);
    }
}
```

**Alternatives Considered**:
- Unlimited context: Rejected - would exhaust LLM context window and memory on long sessions
- Database persistence: Rejected - spec requires session-only, adds unnecessary complexity
- Summary-based compression: Rejected - adds complexity and risks losing important context

---

### Q3: How to render Excel data as formatted HTML tables in Blazor?

**Decision**: Razor component with dynamic table generation from MCP data

**Rationale**:
- Blazor's Razor syntax naturally expresses HTML table structure
- MCP tools return structured JSON data that maps to table rows/columns
- Component-based approach enables reusability and testing
- CSS styling provides borders, headers, alignment per FR-005a
- Server-side rendering ensures compatibility without JavaScript dependency

**Implementation Pattern**:
```razor
@* DataTable.razor *@
<table class="excel-data-table">
    <thead>
        <tr>
            @foreach (var column in Columns)
            {
                <th>@column</th>
            }
        </tr>
    </thead>
    <tbody>
        @foreach (var row in Rows)
        {
            <tr>
                @foreach (var cell in row)
                {
                    <td>@cell</td>
                }
            </tr>
        }
    </tbody>
</table>

@code {
    [Parameter] public List<string> Columns { get; set; }
    [Parameter] public List<List<string>> Rows { get; set; }
}
```

**Alternatives Considered**:
- Markdown tables: Rejected - harder to style, less control over formatting
- JSON display: Rejected - not user-friendly per spec requirements
- Third-party grid component: Rejected - adds dependency, overkill for read-only display

---

### Q4: How to handle multi-workbook session with history markers?

**Decision**: Session state with workbook context tracking and UI separator component

**Rationale**:
- Blazor's component state management handles session scoping
- Store current workbook path and metadata in session state
- When new workbook loaded, insert visual separator in conversation history
- Update session context to point to new workbook
- Agent checks session context before processing queries
- Full history preserved for user reference

**Implementation Pattern**:
```csharp
public class WorkbookSessionState
{
    public string? CurrentWorkbookPath { get; set; }
    public WorkbookMetadata? CurrentMetadata { get; set; }
    public List<ConversationTurn> History { get; } = new();
    
    public void LoadNewWorkbook(string path, WorkbookMetadata metadata)
    {
        // Insert marker in history
        History.Add(new ConversationTurn(
            "system", 
            $"--- Workbook changed to: {Path.GetFileName(path)} ---"
        ));
        
        CurrentWorkbookPath = path;
        CurrentMetadata = metadata;
    }
}
```

**Alternatives Considered**:
- Clear history on workbook change: Rejected - spec explicitly requires preserving history (FR-015a)
- Multiple parallel workbook contexts: Rejected - spec requires only active workbook is queried
- Persistent cross-session state: Rejected - spec requires session-only storage

---

### Q5: How to sanitize error messages while maintaining troubleshooting capability?

**Decision**: Two-tier logging: sanitized UI messages + detailed local logs

**Rationale**:
- Separate error presentation layer from logging layer
- UI displays generic sanitized messages (e.g., "Sheet not found")
- Full details (sheet name, file path, stack trace) written to local JSON Lines log file
- Correlation IDs link UI errors to log entries for troubleshooting
- Complies with FR-010a privacy requirements
- Structured JSON logging enables log analysis tools

**Implementation Pattern**:
```csharp
public class AgentLogger
{
    private readonly ILogger _logger;
    
    public SanitizedError LogAndSanitize(Exception ex, string correlationId)
    {
        // Full details to log file
        _logger.LogError(ex, 
            "Agent error. CorrelationId: {CorrelationId}", 
            correlationId);
        
        // Sanitized for UI
        return new SanitizedError
        {
            Message = "An error occurred while processing your query.",
            CorrelationId = correlationId,
            Timestamp = DateTime.UtcNow
        };
    }
}
```

**Alternatives Considered**:
- Full error details to UI: Rejected - violates privacy requirements (FR-010a)
- No detailed logging: Rejected - makes troubleshooting impossible
- User-controlled verbosity setting: Rejected - adds complexity, risk of accidental exposure

---

## Best Practices Integration

### Semantic Kernel Agent Patterns (per semantic-kernal-bp.md)

**Prompt Templates**:
- Store versioned prompt templates in `PromptTemplates.cs`
- Include system prompt defining agent's role as Excel analyst
- Template variables: {workbook_name}, {sheets_list}, {user_query}
- Guard against off-topic queries with instruction validation

**Context Window Management**:
- 20-turn window fits within typical 8k token context of small models
- Recent context prioritized for relevance
- Workbook metadata included in system context

**Skill Surface Design**:
- Small, testable plugins (WorkbookStructure, Search, DataRetrieval)
- Explicit input/output schemas using SK function attributes
- No sensitive operations beyond read-only workbook access

**Security and Execution Safety**:
- MCP server already sandboxes file access
- No code execution or shell commands
- Agent limited to calling predefined MCP tools
- Timeout enforcement per SC-006 (30 seconds)

**Testing and Reproducibility**:
- Mock MCP client for unit testing plugins
- Canned responses for behavioral tests
- Integration tests with sample workbooks
- Log model version and prompt templates used

**Observability**:
- JSON Lines format for structured logs
- Log: skill invoked, input params, output, execution time, correlation ID
- Correlation IDs link multi-turn conversations
- Local file storage (no external telemetry)

---

## Technology Stack Summary

| Component | Technology | Version | Purpose |
|-----------|-----------|---------|---------|
| UI Framework | Blazor Server | .NET 9.0 | Interactive web UI with real-time updates |
| Agent Framework | Semantic Kernel | Latest stable | LLM orchestration, planning, plugin management |
| MCP Integration | Existing McpClientHost | Current | JSON-RPC client for MCP server communication |
| Excel Processing | ClosedXML | Existing | .xlsx file parsing via MCP server |
| Local LLM Client | Existing LlmStudioClient | Current | OpenAI-compatible API for LM Studio/Ollama |
| Logging | Serilog | Latest | Structured JSON Lines logging |
| Testing | xUnit | Latest | Unit, integration, contract tests |
| Deployment | .NET Publish | Single-file | Self-contained executables per platform |

---

## Risk Mitigation

| Risk | Mitigation |
|------|------------|
| Small LLM struggles with complex queries | Provide clear prompt templates, limit query complexity, use structured tool outputs |
| Context window overflow on long sessions | 20-turn rolling window prevents unbounded growth |
| Blazor Server connection loss | Implement reconnection UI, preserve session state server-side |
| MCP server unresponsive | Retry logic, timeout enforcement, error display per FR-021 |
| Large workbook performance | Stream data, pagination for large result sets, enforce 30s timeout |
| Memory constraints | Rolling window, no persistent history, efficient data structures |

---

## Dependencies & Prerequisites

**Runtime**:
- .NET 9.0 SDK
- Local LLM server (LM Studio, Ollama) with compatible model (<10GB)

**Development**:
- Visual Studio 2022 17.8+ or VS Code with C# DevKit
- PowerShell 7+ for build scripts

**Libraries** (add to ExcelMcp.ChatWeb.csproj):
```xml
<ItemGroup>
  <PackageReference Include="Microsoft.SemanticKernel" Version="1.x.x" />
  <PackageReference Include="Serilog.AspNetCore" Version="8.x.x" />
  <PackageReference Include="Serilog.Sinks.File" Version="5.x.x" />
  <PackageReference Include="Serilog.Formatting.Compact" Version="2.x.x" />
</ItemGroup>
```

---

## Next Steps

Proceed to **Phase 1**: Data Model & Contracts design
- Define ConversationTurn, WorkbookContext, AgentResponse models
- Document SK plugin contracts (input/output schemas)
- Map MCP tools to SK plugin methods
- Create conversation API for Blazor components
