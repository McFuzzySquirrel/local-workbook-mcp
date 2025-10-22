# Quickstart Guide: Local Excel Conversational Agent

**Feature**: 001-local-excel-chat-agent  
**Last Updated**: October 22, 2025

## Overview

This guide helps developers get started implementing the Local Excel Conversational Agent feature. Follow these steps to set up your development environment, understand the architecture, and begin development.

---

## Prerequisites

### Required Software

- **.NET 9.0 SDK** or later
- **Visual Studio 2022 17.8+** or **VS Code** with C# DevKit extension
- **PowerShell 7+** for build scripts
- **Local LLM Server**: LM Studio, Ollama, or compatible OpenAI-like endpoint

### Recommended LLM Models

Small models (<10GB) with function calling support:
- **Phi-4** (Microsoft, ~7GB, excellent for reasoning)
- **Mistral 7B** (Good balance of size and capability)
- **Llama 3.2 8B** (Meta, strong instruction following)

### LM Studio Setup

1. Download from [lmstudio.ai](https://lmstudio.ai/)
2. Install and launch
3. Download a compatible model (e.g., phi-4-mini-reasoning)
4. Start local server on `http://localhost:1234`
5. Note the model identifier for configuration

---

## Project Structure

```
src/ExcelMcp.ChatWeb/          # Main project to enhance
├── Components/                 # NEW - Blazor UI components
│   ├── Pages/Chat.razor       # Main chat page
│   └── Shared/                # Reusable components
├── Services/                   # Agent and business logic
│   ├── Agent/                 # NEW - SK agent services
│   └── Plugins/               # NEW - SK plugins
├── Models/                     # Data models
└── Program.cs                  # Service registration

specs/001-local-excel-chat-agent/  # This feature's docs
├── spec.md                     # Requirements specification
├── plan.md                     # This implementation plan
├── research.md                 # Technology decisions
├── data-model.md              # Entity definitions
├── contracts/                  # API contracts
└── quickstart.md              # This file
```

---

## Development Setup

### 1. Clone and Branch

```powershell
# Ensure you're on the feature branch
git checkout 001-local-excel-chat-agent

# Verify current branch
git branch --show-current
# Should output: 001-local-excel-chat-agent
```

### 2. Install Dependencies

Add required NuGet packages to `ExcelMcp.ChatWeb.csproj`:

```powershell
cd src/ExcelMcp.ChatWeb

# Semantic Kernel
dotnet add package Microsoft.SemanticKernel

# Logging
dotnet add package Serilog.AspNetCore
dotnet add package Serilog.Sinks.File
dotnet add package Serilog.Formatting.Compact

# Existing packages should already be present:
# - ClosedXML (via ExcelMcp.Contracts)
# - System.Text.Json
```

### 3. Configure Local LLM

Update `appsettings.Development.json`:

```json
{
  "LlmStudio": {
    "BaseUrl": "http://localhost:1234",
    "Model": "phi-4-mini-reasoning",
    "MaxTokens": 2048,
    "Temperature": 0.7
  },
  "SemanticKernel": {
    "TimeoutSeconds": 30,
    "MaxRetries": 2
  },
  "Conversation": {
    "MaxContextTurns": 20,
    "MaxResponseLength": 10000
  },
  "Serilog": {
    "MinimumLevel": "Information",
    "WriteTo": [
      {
        "Name": "File",
        "Args": {
          "path": "logs/agent-.log",
          "rollingInterval": "Day",
          "formatter": "Serilog.Formatting.Compact.CompactJsonFormatter, Serilog.Formatting.Compact"
        }
      }
    ]
  }
}
```

### 4. Verify Existing Components

Ensure these existing components are working:

```powershell
# Build solution
dotnet build ExcelLocalMcp.sln

# Run tests
dotnet test

# Start MCP server (test existing functionality)
dotnet run --project src/ExcelMcp.Server -- --workbook "path/to/test.xlsx"
```

---

## Development Workflow

### Phase 1: Core Data Models (Week 1)

**Objective**: Implement data models from `data-model.md`

**Files to Create**:
```
src/ExcelMcp.ChatWeb/Models/
├── ConversationTurn.cs
├── WorkbookContext.cs
├── WorkbookSession.cs
├── AgentResponse.cs
├── ToolInvocation.cs
├── SanitizedError.cs
├── ContentType.cs
└── ErrorCode.cs
```

**Key Points**:
- Use records for immutable data structures where appropriate
- Add validation attributes (Required, Range, etc.)
- Include XML documentation comments
- Follow existing project naming conventions

**Test-First Approach**:
1. Write model tests first (validation, serialization)
2. Implement models to pass tests
3. Commit with test evidence

**Example**:
```csharp
// ConversationTurn.cs
public record ConversationTurn
{
    public Guid Id { get; init; } = Guid.NewGuid();
    
    [Required]
    public required string Role { get; init; }
    
    [Required]
    public required string Content { get; init; }
    
    public DateTimeOffset Timestamp { get; init; } = DateTimeOffset.UtcNow;
    
    [Required]
    public required string CorrelationId { get; init; }
    
    public ContentType ContentType { get; init; } = ContentType.Text;
    
    public Dictionary<string, object>? Metadata { get; init; }
}

// ConversationTurnTests.cs
public class ConversationTurnTests
{
    [Fact]
    public void ConversationTurn_WhenCreated_ShouldHaveValidId()
    {
        // Arrange & Act
        var turn = new ConversationTurn 
        { 
            Role = "user", 
            Content = "test",
            CorrelationId = Guid.NewGuid().ToString()
        };
        
        // Assert
        Assert.NotEqual(Guid.Empty, turn.Id);
    }
    
    [Theory]
    [InlineData("")]
    [InlineData(null)]
    public void ConversationTurn_WhenRoleInvalid_ShouldThrow(string role)
    {
        // Act & Assert
        Assert.Throws<ArgumentNullException>(() => 
            new ConversationTurn 
            { 
                Role = role, 
                Content = "test",
                CorrelationId = "123"
            }
        );
    }
}
```

---

### Phase 2: Semantic Kernel Plugins (Week 2)

**Objective**: Implement SK plugins wrapping MCP tools

**Files to Create**:
```
src/ExcelMcp.ChatWeb/Services/Plugins/
├── WorkbookStructurePlugin.cs
├── WorkbookSearchPlugin.cs
├── DataRetrievalPlugin.cs
└── PluginDescriptors.cs
```

**Key Points**:
- Use `[KernelFunction]` and `[Description]` attributes
- Return JSON strings (SK handles parsing)
- Implement error handling per `mcp-tool-mapping.md`
- Add structured logging for all tool invocations

**Test Strategy**:
1. Mock `IMcpClient` for unit tests
2. Test each plugin method independently
3. Verify correct MCP tool called with right parameters
4. Test error scenarios (MCP failure, timeout, invalid input)

**Example**:
```csharp
// WorkbookStructurePlugin.cs
public class WorkbookStructurePlugin
{
    private readonly IMcpClient _mcpClient;
    private readonly ILogger<WorkbookStructurePlugin> _logger;
    
    public WorkbookStructurePlugin(
        IMcpClient mcpClient, 
        ILogger<WorkbookStructurePlugin> logger)
    {
        _mcpClient = mcpClient;
        _logger = logger;
    }
    
    [KernelFunction("list_workbook_structure")]
    [Description("Lists all sheets and tables in the loaded workbook")]
    [return: Description("JSON string with sheets, tables, and metadata")]
    public async Task<string> ListWorkbookStructure()
    {
        var correlationId = Guid.NewGuid().ToString();
        var sw = Stopwatch.StartNew();
        
        try
        {
            var result = await _mcpClient.CallToolAsync("excel-list-structure");
            
            _logger.LogInformation(
                "MCP tool invoked successfully. Tool: {Tool}, CorrelationId: {CorrelationId}, Duration: {DurationMs}ms",
                "excel-list-structure", correlationId, sw.ElapsedMilliseconds
            );
            
            return JsonSerializer.Serialize(result);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex,
                "MCP tool invocation failed. Tool: {Tool}, CorrelationId: {CorrelationId}",
                "excel-list-structure", correlationId
            );
            
            return JsonSerializer.Serialize(new ErrorResponse
            {
                Error = true,
                ErrorCode = "MCP_ERROR",
                Message = "Could not retrieve workbook structure",
                CorrelationId = correlationId,
                CanRetry = true
            });
        }
    }
}

// WorkbookStructurePluginTests.cs
public class WorkbookStructurePluginTests
{
    private readonly Mock<IMcpClient> _mockMcpClient;
    private readonly Mock<ILogger<WorkbookStructurePlugin>> _mockLogger;
    private readonly WorkbookStructurePlugin _plugin;
    
    public WorkbookStructurePluginTests()
    {
        _mockMcpClient = new Mock<IMcpClient>();
        _mockLogger = new Mock<ILogger<WorkbookStructurePlugin>>();
        _plugin = new WorkbookStructurePlugin(_mockMcpClient.Object, _mockLogger.Object);
    }
    
    [Fact]
    public async Task ListWorkbookStructure_WhenMcpSucceeds_ReturnsStructure()
    {
        // Arrange
        var mcpResult = JsonDocument.Parse(@"{
            ""sheets"": [{""name"": ""Sales"", ""rowCount"": 100}]
        }").RootElement;
        
        _mockMcpClient
            .Setup(m => m.CallToolAsync("excel-list-structure", null, default))
            .ReturnsAsync(mcpResult);
        
        // Act
        var result = await _plugin.ListWorkbookStructure();
        
        // Assert
        Assert.NotNull(result);
        Assert.Contains("Sales", result);
        _mockMcpClient.Verify(
            m => m.CallToolAsync("excel-list-structure", null, default), 
            Times.Once
        );
    }
}
```

---

### Phase 3: Conversation Management (Week 2)

**Objective**: Implement conversation history and rolling window

**Files to Create**:
```
src/ExcelMcp.ChatWeb/Services/Agent/
├── ConversationManager.cs
├── IConversationManager.cs
└── ConversationOptions.cs
```

**Key Features**:
- 20-turn rolling window using SK `ChatHistory`
- Full history storage for UI display
- System message handling
- Workbook switch markers

**Example**:
```csharp
public class ConversationManager : IConversationManager
{
    private readonly ChatHistory _contextWindow = new();
    private readonly List<ConversationTurn> _fullHistory = new();
    private readonly IOptions<ConversationOptions> _options;
    private readonly ILogger<ConversationManager> _logger;
    
    private const int MaxContextTurns = 20;
    
    public void AddUserTurn(string message, string correlationId)
    {
        var turn = new ConversationTurn
        {
            Role = "user",
            Content = message,
            CorrelationId = correlationId,
            ContentType = ContentType.Text
        };
        
        _fullHistory.Add(turn);
        _contextWindow.AddUserMessage(message);
        EnforceWindowLimit();
        
        _logger.LogDebug(
            "User turn added. CorrelationId: {CorrelationId}, ContextWindowSize: {Size}",
            correlationId, _contextWindow.Count
        );
    }
    
    private void EnforceWindowLimit()
    {
        while (_contextWindow.Count > MaxContextTurns)
        {
            _contextWindow.RemoveAt(0);
            _logger.LogDebug("Evicted oldest turn from context window");
        }
    }
    
    public ChatHistory GetContextForLLM() => _contextWindow;
    
    public List<ConversationTurn> GetFullHistory() => _fullHistory.ToList();
}
```

---

### Phase 4: Agent Service (Week 3)

**Objective**: Implement main ExcelAgentService orchestrating SK and plugins

**Files to Create**:
```
src/ExcelMcp.ChatWeb/Services/Agent/
├── ExcelAgentService.cs
├── IExcelAgentService.cs
├── ResponseFormatter.cs
├── IResponseFormatter.cs
└── PromptTemplates.cs
```

**Key Features**:
- Query processing with SK
- Plugin invocation
- Response formatting (HTML tables)
- Error sanitization
- Timeout enforcement

**Example** (simplified):
```csharp
public class ExcelAgentService : IExcelAgentService
{
    private readonly Kernel _kernel;
    private readonly IConversationManager _conversationManager;
    private readonly IResponseFormatter _formatter;
    private readonly ILogger<ExcelAgentService> _logger;
    
    public async Task<AgentResponse> ProcessQueryAsync(
        string query, 
        WorkbookSession session,
        CancellationToken cancellationToken = default)
    {
        var correlationId = Guid.NewGuid().ToString();
        var sw = Stopwatch.StartNew();
        
        try
        {
            // Add user turn
            _conversationManager.AddUserTurn(query, correlationId);
            
            // Get context for LLM
            var chatHistory = _conversationManager.GetContextForLLM();
            
            // Invoke SK with timeout
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(TimeSpan.FromSeconds(30));
            
            var result = await _kernel.InvokePromptAsync(
                PromptTemplates.ChatPrompt,
                new KernelArguments 
                { 
                    ["chat_history"] = chatHistory,
                    ["workbook_name"] = session.CurrentContext?.WorkbookName ?? "No workbook loaded"
                },
                cancellationToken: cts.Token
            );
            
            // Format response
            var response = _formatter.FormatResponse(result.ToString(), correlationId);
            
            // Add assistant turn
            _conversationManager.AddAssistantTurn(
                response.Content, 
                correlationId, 
                response.ContentType
            );
            
            response.ProcessingTimeMs = (int)sw.ElapsedMilliseconds;
            
            return response;
        }
        catch (OperationCanceledException)
        {
            _logger.LogWarning("Query timeout. CorrelationId: {CorrelationId}", correlationId);
            return _formatter.CreateTimeoutError(correlationId);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Query processing failed. CorrelationId: {CorrelationId}", correlationId);
            return _formatter.SanitizeError(ex, correlationId);
        }
    }
}
```

---

### Phase 5: Blazor UI Components (Week 4)

**Objective**: Build interactive chat interface

**Files to Create**:
```
src/ExcelMcp.ChatWeb/Components/
├── Pages/
│   ├── Chat.razor
│   └── Chat.razor.cs
├── Shared/
│   ├── ChatMessage.razor
│   ├── DataTable.razor
│   ├── WorkbookSelector.razor
│   └── LoadingIndicator.razor
└── Layout/
    └── MainLayout.razor (modify existing)
```

**Example Chat.razor**:
```razor
@page "/chat"
@inject IExcelAgentService AgentService
@inject WorkbookSession Session

<PageTitle>Excel Chat Agent</PageTitle>

<div class="chat-container">
    <div class="chat-header">
        <WorkbookSelector OnWorkbookSelected="HandleWorkbookSelected" />
        <button @onclick="ClearHistory">Clear History</button>
    </div>
    
    <div class="chat-messages">
        @foreach (var turn in _conversationHistory)
        {
            <ChatMessage Turn="turn" />
        }
        
        @if (_isProcessing)
        {
            <LoadingIndicator Message="Processing query..." />
        }
    </div>
    
    <div class="chat-input">
        <textarea @bind="_currentQuery" 
                  @onkeydown="HandleKeyDown"
                  placeholder="Ask a question about your workbook..."
                  disabled="@_isProcessing">
        </textarea>
        <button @onclick="SubmitQuery" disabled="@(_isProcessing || string.IsNullOrWhiteSpace(_currentQuery))">
            Send
        </button>
    </div>
</div>

@code {
    private string _currentQuery = "";
    private bool _isProcessing = false;
    private List<ConversationTurn> _conversationHistory = new();
    
    private async Task SubmitQuery()
    {
        if (string.IsNullOrWhiteSpace(_currentQuery)) return;
        
        _isProcessing = true;
        var query = _currentQuery;
        _currentQuery = "";
        
        try
        {
            var response = await AgentService.ProcessQueryAsync(query, Session);
            _conversationHistory = Session.ConversationHistory;
        }
        finally
        {
            _isProcessing = false;
        }
    }
}
```

---

## Testing Strategy

### Unit Tests

Test each component in isolation:
- Models: Validation, serialization
- Plugins: MCP interaction, error handling
- Services: Business logic, orchestration
- Formatters: HTML generation, sanitization

### Integration Tests

Test component interaction:
- Agent service + plugins + MCP client
- Conversation manager + SK ChatHistory
- End-to-end query flow

### Contract Tests

Verify API contracts:
- SK plugin signatures match specifications
- MCP tool calls use correct parameters
- Response formats match data model

---

## Running the Application

### Start LM Studio

1. Launch LM Studio
2. Load your chosen model
3. Start local server (default: `localhost:1234`)

### Run the Application

```powershell
# From repository root
dotnet run --project src/ExcelMcp.ChatWeb

# Application starts on http://localhost:5000
# Navigate to http://localhost:5000/chat
```

### Test with Sample Workbook

1. Prepare a test .xlsx file with multiple sheets
2. Click "Load Workbook" in UI
3. Select your test file
4. Try queries:
   - "What sheets are in this workbook?"
   - "Show me the first 10 rows of the Sales table"
   - "Search for 'Project X' across all sheets"

---

## Troubleshooting

### LLM Not Responding

- Verify LM Studio is running and model loaded
- Check `appsettings.Development.json` BaseUrl matches LM Studio
- Look in logs for connection errors

### MCP Server Issues

- Ensure workbook file exists and is valid .xlsx
- Check ExcelMcp.Server logs
- Verify file permissions

### Blazor Circuit Disconnects

- Check browser console for errors
- Verify WebSocket connection
- Check server logs for exceptions

### Performance Issues

- Reduce conversation context window size
- Use smaller workbooks for testing
- Check LLM model size and hardware capabilities

---

## Next Steps After Quickstart

1. **Read Detailed Specs**: Review `spec.md`, `data-model.md`, and contract docs
2. **Follow Test-First**: Write tests before implementation
3. **Incremental Development**: Implement one user story at a time (P1 → P4)
4. **Log Everything**: Use structured logging for debugging
5. **Performance Monitoring**: Track query times, token usage

## Resources

- **Semantic Kernel Docs**: [learn.microsoft.com/semantic-kernel](https://learn.microsoft.com/semantic-kernel)
- **Blazor Docs**: [learn.microsoft.com/aspnet/core/blazor](https://learn.microsoft.com/aspnet/core/blazor)
- **MCP Spec**: [modelcontextprotocol.io](https://modelcontextprotocol.io)
- **ClosedXML Docs**: [github.com/ClosedXML/ClosedXML](https://github.com/ClosedXML/ClosedXML)

---

**Questions?** Check the spec documentation or create an issue in the repository.
