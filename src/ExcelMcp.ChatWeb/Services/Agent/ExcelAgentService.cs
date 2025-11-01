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

            // Initialize/reinitialize MCP client with this workbook
            if (_mcpClient is McpClientHost mcpHost)
            {
                await mcpHost.InitializeWithWorkbookAsync(filePath, null, cancellationToken);
            }

            // Call MCP to get workbook structure
            var result = await _mcpClient.CallToolAsync("excel-list-structure", null, cancellationToken);
            
            if (result.IsError || result.Content == null || result.Content.Count == 0)
            {
                throw new InvalidOperationException("Failed to retrieve workbook structure from MCP server");
            }

            // Extract the text content from the first content item (should be JSON)
            var firstContent = result.Content[0];
            var metadataJson = firstContent.Text;
            
            if (string.IsNullOrEmpty(metadataJson))
            {
                throw new InvalidOperationException("Empty response from MCP server");
            }

            // Deserialize the WorkbookMetadata directly from the MCP response
            var jsonOptions = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase
            };
            var metadata = JsonSerializer.Deserialize<WorkbookMetadata>(metadataJson, jsonOptions);
            
            if (metadata == null)
            {
                throw new InvalidOperationException("Failed to deserialize workbook metadata");
            }

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

            // Create timeout cancellation token from config
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(TimeSpan.FromSeconds(_skOptions.TimeoutSeconds));

            // Get conversation context for LLM
            var chatHistory = _conversationManager.GetContextForLLM();

            // Add comprehensive system prompt (workbook-agnostic, proven from CLI agent)
            chatHistory.AddSystemMessage(PromptTemplates.GetCompleteSystemPrompt());
            
            // Add specific context about current workbook
            var sheets = session.CurrentContext.Metadata?.Worksheets.Select(w => w.Name).ToList() ?? new List<string>();
            var tables = session.CurrentContext.Metadata?.Worksheets
                .SelectMany(w => w.Tables.Select(t => t.Name))
                .ToList() ?? new List<string>();

            var workbookContext = $@"CURRENT WORKBOOK CONTEXT:
- File: '{session.CurrentContext.WorkbookName}'
- Sheets ({sheets.Count}): {string.Join(", ", sheets.Take(10))}{(sheets.Count > 10 ? "..." : "")}
- Tables: {(tables.Any() ? string.Join(", ", tables.Take(10)) : "None defined")}{(tables.Count > 10 ? "..." : "")}

Remember: Use tools to discover structure if you need more details!";
            
            chatHistory.AddSystemMessage(workbookContext);

            // Invoke Semantic Kernel with automatic function calling enabled
            var chatService = _kernel.GetRequiredService<IChatCompletionService>();
            
            // Configure execution settings to enable automatic function calling
            var executionSettings = new OpenAIPromptExecutionSettings
            {
                ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions,
                Temperature = 0.7,
                MaxTokens = 2048
            };

            ChatMessageContent result;
            try
            {
                result = await chatService.GetChatMessageContentAsync(
                    chatHistory,
                    executionSettings: executionSettings,
                    kernel: _kernel,
                    cancellationToken: timeoutCts.Token);
            }
            catch (Exception ex) when (
                (ex is NullReferenceException || ex is ArgumentOutOfRangeException) && 
                (ex.StackTrace?.Contains("get_Refusal") ?? false))
            {
                // Known issue: OpenAI SDK expects a 'refusal' field that LM Studio doesn't return
                // Provide helpful error message
                _conversationManager.AddSystemMessage(
                    "❌ Error: Your LM Studio server is not returning OpenAI-compatible responses.\n\n" +
                    "**To fix this:**\n" +
                    "1. In LM Studio, go to the **Developer** tab\n" +
                    "2. Look for **'Response format'** or **'OpenAI compatibility'** settings\n" +
                    "3. Enable full OpenAI API compatibility\n" +
                    "4. Restart the LM Studio server\n\n" +
                    "Alternatively, try using a different model or updating LM Studio to the latest version.");
                
                throw new InvalidOperationException(
                    "LM Studio compatibility issue - see chat for details", ex);
            }

            var responseContent = result.Content ?? "No response generated.";
            var duration = DateTimeOffset.UtcNow - startTime;

            // Detect and format table data
            var (formattedContent, contentType) = FormatResponseContent(responseContent);

            // Add assistant response to conversation
            _conversationManager.AddAssistantTurn(formattedContent, correlationId, contentType);

            // Build agent response
            var response = new AgentResponse
            {
                CorrelationId = correlationId,
                Content = formattedContent,
                ContentType = contentType,
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
            // Timeout
            var timeoutMessage = $"Query timed out after {_skOptions.TimeoutSeconds} seconds";
            _logger.LogError(correlationId, new TimeoutException("Query timeout"), timeoutMessage);
            return CreateErrorResponse(correlationId, new TimeoutException(timeoutMessage));
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
    /// Formats response content appropriately based on content type detection.
    /// Handles CSV data, HTML tables, and plain text.
    /// Returns formatted content and appropriate content type.
    /// </summary>
    private (string content, ContentType type) FormatResponseContent(string responseContent)
    {
        if (string.IsNullOrWhiteSpace(responseContent))
        {
            return (string.Empty, ContentType.Text);
        }

        // Check if response already contains HTML table (LLM may generate this)
        if (responseContent.Contains("<table") && responseContent.Contains("</table>"))
        {
            // Already has HTML table - just ensure it's properly wrapped
            if (!responseContent.Contains("class=\"data-table\""))
            {
                responseContent = responseContent.Replace("<table", "<table class=\"data-table\"");
            }
            return (responseContent, ContentType.Table);
        }

        // Try to detect and parse CSV/tabular data from tool responses
        if (TryParseCsvToTable(responseContent, out var tableData))
        {
            var htmlTable = _responseFormatter.FormatAsHtmlTable(tableData);
            
            // If LLM added commentary before/after the CSV, preserve it
            var tableStartIndex = FindTableDataStart(responseContent);
            if (tableStartIndex > 0)
            {
                var commentary = responseContent.Substring(0, tableStartIndex).Trim();
                if (!string.IsNullOrWhiteSpace(commentary))
                {
                    var formattedCommentary = _responseFormatter.FormatAsText(commentary);
                    return ($"{formattedCommentary}\n\n{htmlTable}", ContentType.Table);
                }
            }
            
            return (htmlTable, ContentType.Table);
        }

        // No table detected, return as formatted text
        var formattedText = _responseFormatter.FormatAsText(responseContent);
        return (formattedText, ContentType.Text);
    }

    /// <summary>
    /// Finds where tabular data starts in a response (after any introductory text).
    /// </summary>
    private static int FindTableDataStart(string content)
    {
        var lines = content.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            // Look for CSV header pattern (multiple comma-separated words)
            if (line.Split(',').Length >= 2 && 
                line.Split(',').All(part => !string.IsNullOrWhiteSpace(part)))
            {
                return content.IndexOf(line);
            }
        }
        return 0;
    }

    /// <summary>
    /// Attempts to parse CSV data from response content into TableData model.
    /// </summary>
    private static bool TryParseCsvToTable(string content, out TableData tableData)
    {
        tableData = new TableData { Columns = new List<string>(), Rows = new List<List<string>>() };

        // Look for CSV patterns: multiple lines with consistent comma delimiters
        var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        
        if (lines.Length < 2)
        {
            return false; // Need at least header + 1 data row
        }

        // Parse first line as headers
        var headers = lines[0].Split(',').Select(h => h.Trim()).ToList();
        
        if (headers.Count < 2)
        {
            return false; // Need at least 2 columns to be a table
        }

        // Verify subsequent lines have same number of columns
        var rows = new List<List<string>>();
        for (int i = 1; i < lines.Length; i++)
        {
            var cells = lines[i].Split(',').Select(c => c.Trim()).ToList();
            
            // Allow some flexibility (within 1 column difference)
            if (Math.Abs(cells.Count - headers.Count) > 1)
            {
                return false; // Not consistent column structure
            }

            // Pad if needed
            while (cells.Count < headers.Count)
            {
                cells.Add(string.Empty);
            }

            // Truncate if too many
            if (cells.Count > headers.Count)
            {
                cells = cells.Take(headers.Count).ToList();
            }

            rows.Add(cells);
        }

        // Success - we have table data
        tableData.Columns = headers;
        tableData.Rows = rows;
        tableData.Metadata = new Models.TableMetadata
        {
            RowCount = rows.Count,
            IsTruncated = false
        };

        return true;
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
