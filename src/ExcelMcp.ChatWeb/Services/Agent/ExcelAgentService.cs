using System.Text.Json;
using ExcelMcp.ChatWeb.Logging;
using ExcelMcp.ChatWeb.Models;
using ExcelMcp.ChatWeb.Options;
using ExcelMcp.Contracts;
using Microsoft.Extensions.Options;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;

namespace ExcelMcp.ChatWeb.Services.Agent;

/// <summary>
/// Core agent service orchestrating LLM interactions with Excel workbooks.
/// Handles query processing, workbook loading, and conversation management.
/// </summary>
public class ExcelAgentService : IExcelAgentService
{
    private readonly Kernel _kernel;
    private readonly IConversationManager _conversationManager;
    private readonly IResponseFormatter _responseFormatter;
    private readonly IMcpClient _mcpClient;
    private readonly AgentLogger _logger;
    private readonly SemanticKernelOptions _skOptions;

    public ExcelAgentService(
        Kernel kernel,
        IConversationManager conversationManager,
        IResponseFormatter responseFormatter,
        IMcpClient mcpClient,
        AgentLogger logger,
        IOptions<SemanticKernelOptions> skOptions)
    {
        _kernel = kernel ?? throw new ArgumentNullException(nameof(kernel));
        _conversationManager = conversationManager ?? throw new ArgumentNullException(nameof(conversationManager));
        _responseFormatter = responseFormatter ?? throw new ArgumentNullException(nameof(responseFormatter));
        _mcpClient = mcpClient ?? throw new ArgumentNullException(nameof(mcpClient));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        _skOptions = skOptions?.Value ?? throw new ArgumentNullException(nameof(skOptions));
    }

    /// <summary>
    /// Loads a workbook, retrieves metadata via MCP, and updates session context.
    /// </summary>
    public async Task<WorkbookContext> LoadWorkbookAsync(
        string filePath,
        WorkbookSession session,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path cannot be empty", nameof(filePath));
        }

        if (session == null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        var correlationId = Guid.NewGuid().ToString("N")[..12];
        _logger.LogQuery(correlationId, $"Loading workbook: {filePath}");

        try
        {
            // Validate file exists
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Workbook file not found: {filePath}");
            }

            // Validate file extension
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension != ".xlsx" && extension != ".xls")
            {
                throw new ArgumentException($"Invalid file type. Expected .xlsx or .xls, got {extension}");
            }

            // Call MCP to get workbook structure
            var result = await _mcpClient.CallToolAsync("excel-list-structure", null, cancellationToken);
            
            if (result.IsError || result.Content == null)
            {
                throw new InvalidOperationException("Failed to retrieve workbook structure from MCP server");
            }

            var structureJson = result.Content.ToString();
            if (string.IsNullOrEmpty(structureJson))
            {
                throw new InvalidOperationException("Empty response from MCP server");
            }

            var structure = JsonSerializer.Deserialize<JsonDocument>(structureJson);
            var sheetsElement = structure?.RootElement.GetProperty("sheets");
            
            var worksheets = new List<WorksheetMetadata>();

            if (sheetsElement.HasValue)
            {
                foreach (var sheet in sheetsElement.Value.EnumerateArray())
                {
                    var sheetName = sheet.GetProperty("name").GetString() ?? "Unknown";
                    var tables = new List<ExcelMcp.Contracts.TableMetadata>();
                    var columnHeaders = new List<string>();

                    if (sheet.TryGetProperty("tables", out var tablesElement))
                    {
                        foreach (var table in tablesElement.EnumerateArray())
                        {
                            var tableName = table.GetString();
                            if (!string.IsNullOrEmpty(tableName))
                            {
                                // Create table metadata (we'll get details later if needed)
                                tables.Add(new ExcelMcp.Contracts.TableMetadata(
                                    tableName,
                                    sheetName,
                                    new List<string>(), // Column headers not available from list-structure
                                    0 // Row count not available
                                ));
                            }
                        }
                    }

                    worksheets.Add(new WorksheetMetadata(sheetName, tables, columnHeaders));
                }
            }

            var metadata = new WorkbookMetadata(
                filePath,
                worksheets,
                DateTimeOffset.UtcNow
            );

            var context = new WorkbookContext
            {
                WorkbookPath = filePath,
                WorkbookName = Path.GetFileName(filePath),
                Metadata = metadata,
                LoadedAt = DateTimeOffset.UtcNow,
                IsValid = true
            };

            // Update session
            session.LoadNewWorkbook(context);

            // Add system message about workbook change
            _conversationManager.AddSystemMessage($"Workbook changed to {context.WorkbookName}");

            return context;
        }
        catch (Exception ex)
        {
            _logger.LogError(correlationId, ex, "Failed to load workbook");
            
            return new WorkbookContext
            {
                WorkbookPath = filePath,
                WorkbookName = Path.GetFileName(filePath),
                IsValid = false,
                LoadedAt = DateTimeOffset.UtcNow,
                ErrorMessage = ex.Message
            };
        }
    }

    /// <summary>
    /// Processes a user query using Semantic Kernel with conversation context.
    /// </summary>
    public async Task<AgentResponse> ProcessQueryAsync(
        string query,
        WorkbookSession session,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(query))
        {
            throw new ArgumentException("Query cannot be empty", nameof(query));
        }

        if (session == null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        var correlationId = Guid.NewGuid().ToString("N")[..12];
        _logger.LogQuery(correlationId, query);

        var startTime = DateTimeOffset.UtcNow;

        try
        {
            // Check if workbook is loaded
            if (session.CurrentContext == null || !session.CurrentContext.IsValid)
            {
                throw new InvalidOperationException("No valid workbook is currently loaded. Please load a workbook first.");
            }

            // Add user message to conversation
            _conversationManager.AddUserTurn(query, correlationId);

            // Create timeout cancellation token (30 seconds)
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(30));

            // Get conversation context for LLM
            var chatHistory = _conversationManager.GetContextForLLM();

            // Add system context about current workbook
            var sheets = session.CurrentContext.Metadata?.Worksheets.Select(w => w.Name).ToList() ?? new List<string>();
            var tables = session.CurrentContext.Metadata?.Worksheets
                .SelectMany(w => w.Tables.Select(t => t.Name))
                .ToList() ?? new List<string>();

            var systemContext = $@"You are analyzing the workbook '{session.CurrentContext.WorkbookName}' 
with {sheets.Count} sheets: {string.Join(", ", sheets)}. 
Available tables: {string.Join(", ", tables)}.

IMPORTANT: When you use preview_table and receive CSV data, output it as an HTML table using this format:
<table class='data-table'>
<thead><tr><th>Column1</th><th>Column2</th></tr></thead>
<tbody>
<tr><td>value1</td><td>value2</td></tr>
</tbody>
</table>

Do NOT describe the data or reformat it as markdown. Present the actual data in HTML table format.";
            
            chatHistory.AddSystemMessage(systemContext);

            // Invoke Semantic Kernel with automatic function calling enabled
            var chatService = _kernel.GetRequiredService<IChatCompletionService>();
            
            // Configure execution settings to enable automatic function calling
            var executionSettings = new OpenAIPromptExecutionSettings
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions,
                Temperature = 0.7,
                MaxTokens = 2048
            };

            var result = await chatService.GetChatMessageContentAsync(
                chatHistory,
                executionSettings: executionSettings,
                kernel: _kernel,
                cancellationToken: timeoutCts.Token);

            var responseContent = result.Content ?? "No response generated.";
            var duration = DateTimeOffset.UtcNow - startTime;

            // Add assistant response to conversation
            _conversationManager.AddAssistantTurn(responseContent, correlationId, ContentType.Text);

            // Build agent response
            var response = new AgentResponse
            {
                CorrelationId = correlationId,
                Content = responseContent,
                ContentType = ContentType.Text,
                ProcessingTimeMs = (int)duration.TotalMilliseconds,
                ModelUsed = _skOptions.Model
            };

            _logger.LogResponse(response);

            return response;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // User-initiated cancellation
            _logger.LogError(correlationId, new OperationCanceledException("Query cancelled by user"), "Query cancelled");
            return CreateErrorResponse(correlationId, new OperationCanceledException("Query cancelled by user"));
        }
        catch (OperationCanceledException)
        {
            // Timeout (30 seconds)
            _logger.LogError(correlationId, new TimeoutException("Query timeout"), "Query timed out after 30 seconds");
            return CreateErrorResponse(correlationId, new TimeoutException("Query timed out after 30 seconds"));
        }
        catch (Exception ex)
        {
            _logger.LogError(correlationId, ex, "Error processing query");
            return CreateErrorResponse(correlationId, ex);
        }
    }

    /// <summary>
    /// Clears conversation history while keeping the workbook loaded.
    /// </summary>
    public Task ClearConversationAsync(WorkbookSession session)
    {
        if (session == null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        // Clear the conversation history in the session
        session.ConversationHistory.Clear();
        session.ContextWindow.Clear();
        
        // Add system message if workbook is still loaded
        if (session.CurrentContext != null && session.CurrentContext.IsValid)
        {
            _conversationManager.AddSystemMessage($"Conversation cleared. Workbook '{session.CurrentContext.WorkbookName}' is still loaded.");
        }

        return Task.CompletedTask;
    }

    /// <summary>
    /// Generates suggested follow-up queries based on workbook metadata and conversation history.
    /// </summary>
    public Task<List<string>> GetSuggestedQueriesAsync(
        WorkbookSession session,
        int maxSuggestions = 3)
    {
        if (session == null)
        {
            throw new ArgumentNullException(nameof(session));
        }

        var suggestions = new List<string>();

        if (session.CurrentContext == null || !session.CurrentContext.IsValid)
        {
            return Task.FromResult(suggestions);
        }

        var workbook = session.CurrentContext;
        var history = session.ConversationHistory;
        var sheets = workbook.Metadata?.Worksheets.Select(w => w.Name).ToList() ?? new List<string>();
        var tables = workbook.Metadata?.Worksheets
            .SelectMany(w => w.Tables.Select(t => t.Name))
            .ToList() ?? new List<string>();

        // Generate suggestions based on workbook structure
        if (history.Count == 0 || history.All(t => t.Role == "system"))
        {
            // First-time suggestions (no conversation yet)
            if (sheets.Any())
            {
                suggestions.Add($"What sheets are in this workbook?");
            }
            if (tables.Any())
            {
                suggestions.Add($"Show me the structure of the {tables.First()} table");
            }
            if (sheets.Any())
            {
                suggestions.Add($"What data is in the {sheets.First()} sheet?");
            }
        }
        else
        {
            // Context-aware suggestions based on recent conversation
            var recentQuery = history.LastOrDefault(t => t.Role == "user")?.Content?.ToLowerInvariant();

            if (recentQuery?.Contains("structure") == true || recentQuery?.Contains("sheets") == true)
            {
                suggestions.Add("Show me the first 10 rows of data");
                suggestions.Add("Search for specific values in the workbook");
            }
            else if (recentQuery?.Contains("search") == true)
            {
                suggestions.Add("Show me more details about those results");
                suggestions.Add("Can you summarize this data?");
            }
            else if (tables.Any())
            {
                suggestions.Add($"Calculate the sum of a column in {tables.First()}");
                suggestions.Add("Show me statistics for this data");
            }

            // Generic helpful suggestions
            if (suggestions.Count < maxSuggestions)
            {
                suggestions.Add("What other sheets can I explore?");
            }
        }

        return Task.FromResult(suggestions.Take(maxSuggestions).ToList());
    }

    /// <summary>
    /// Creates an error response with sanitized error information.
    /// </summary>
    private AgentResponse CreateErrorResponse(string correlationId, Exception exception)
    {
        var sanitizedError = _responseFormatter.SanitizeErrorMessage(exception, correlationId);

        return new AgentResponse
        {
            CorrelationId = correlationId,
            Content = sanitizedError.Message,
            ContentType = ContentType.Error,
            Error = sanitizedError
        };
    }
}
