using ExcelMcp.ChatWeb.Models;
using Serilog;

namespace ExcelMcp.ChatWeb.Logging;

/// <summary>
/// Structured logging for agent operations with correlation tracking.
/// </summary>
public class AgentLogger
{
    private readonly Serilog.ILogger _logger;

    public AgentLogger(Serilog.ILogger logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Logs a user query with correlation ID.
    /// </summary>
    public void LogQuery(string correlationId, string query, string? workbookName = null)
    {
        _logger.Information(
            "Agent query received. CorrelationId: {CorrelationId}, Workbook: {WorkbookName}, QueryLength: {QueryLength}",
            correlationId,
            workbookName ?? "None",
            query.Length);
    }

    /// <summary>
    /// Logs an MCP tool invocation with timing and success status.
    /// </summary>
    public void LogToolInvocation(ToolInvocation invocation, string correlationId)
    {
        if (invocation.Success)
        {
            _logger.Information(
                "Tool invoked successfully. CorrelationId: {CorrelationId}, Tool: {ToolName}, Plugin: {PluginMethod}, Duration: {DurationMs}ms",
                correlationId,
                invocation.ToolName,
                invocation.PluginMethod,
                invocation.DurationMs);
        }
        else
        {
            _logger.Warning(
                "Tool invocation failed. CorrelationId: {CorrelationId}, Tool: {ToolName}, Plugin: {PluginMethod}, Duration: {DurationMs}ms, Error: {ErrorMessage}",
                correlationId,
                invocation.ToolName,
                invocation.PluginMethod,
                invocation.DurationMs,
                invocation.ErrorMessage);
        }
    }

    /// <summary>
    /// Logs an agent response with processing time and tool count.
    /// </summary>
    public void LogResponse(AgentResponse response)
    {
        _logger.Information(
            "Agent response generated. CorrelationId: {CorrelationId}, ContentType: {ContentType}, ProcessingTime: {ProcessingTimeMs}ms, ToolsInvoked: {ToolCount}, Model: {Model}",
            response.CorrelationId,
            response.ContentType,
            response.ProcessingTimeMs,
            response.ToolsInvoked.Count,
            response.ModelUsed ?? "Unknown");
    }

    /// <summary>
    /// Logs a detailed error with full exception information (not sanitized - for troubleshooting).
    /// </summary>
    public void LogError(string correlationId, Exception exception, string context)
    {
        _logger.Error(
            exception,
            "Agent error occurred. CorrelationId: {CorrelationId}, Context: {Context}, ExceptionType: {ExceptionType}",
            correlationId,
            context,
            exception.GetType().Name);
    }

    /// <summary>
    /// Logs workbook load operation.
    /// </summary>
    public void LogWorkbookLoad(string correlationId, string workbookPath, bool success, int? sheetCount = null)
    {
        if (success)
        {
            _logger.Information(
                "Workbook loaded successfully. CorrelationId: {CorrelationId}, Workbook: {WorkbookName}, Sheets: {SheetCount}",
                correlationId,
                Path.GetFileName(workbookPath),
                sheetCount ?? 0);
        }
        else
        {
            _logger.Warning(
                "Workbook load failed. CorrelationId: {CorrelationId}, WorkbookPath: {WorkbookPath}",
                correlationId,
                workbookPath);
        }
    }

    /// <summary>
    /// Logs conversation clear operation.
    /// </summary>
    public void LogConversationCleared(string sessionId, int turnCount)
    {
        _logger.Information(
            "Conversation cleared. SessionId: {SessionId}, TurnsCleared: {TurnCount}",
            sessionId,
            turnCount);
    }
}
