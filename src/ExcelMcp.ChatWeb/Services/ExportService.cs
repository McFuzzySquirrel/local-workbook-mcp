using System.Text;
using ExcelMcp.ChatWeb.Models;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;

namespace ExcelMcp.ChatWeb.Services;

public class ExportService
{
    private readonly Kernel _kernel;

    public ExportService(Kernel kernel)
    {
        _kernel = kernel;
    }

    /// <summary>
    /// Exports the full conversation history as a Markdown file.
    /// </summary>
    public byte[] ExportConversation(WorkbookSession session)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"# Conversation Export - {session.CurrentContext?.WorkbookName ?? "Unknown Workbook"}");
        sb.AppendLine($"Date: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        foreach (var turn in session.ConversationHistory)
        {
            sb.AppendLine($"## {turn.Role.ToUpperInvariant()} ({turn.Timestamp:HH:mm:ss})");
            sb.AppendLine();
            sb.AppendLine(turn.Content);
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine();
        }

        return Encoding.UTF8.GetBytes(sb.ToString());
    }

    /// <summary>
    /// Exports table data as CSV.
    /// </summary>
    public byte[] ExportDataView(TableData tableData)
    {
        var sb = new StringBuilder();

        // Headers
        sb.AppendLine(string.Join(",", tableData.Columns.Select(EscapeCsv)));

        // Rows
        foreach (var row in tableData.Rows)
        {
            sb.AppendLine(string.Join(",", row.Select(EscapeCsv)));
        }

        return Encoding.UTF8.GetBytes(sb.ToString());
    }

    /// <summary>
    /// Generates a concise summary of key findings using LLM.
    /// </summary>
    public async Task<string> ExportInsightsSummaryAsync(WorkbookSession session)
    {
        if (!session.ConversationHistory.Any()) return "No conversation to summarize.";

        var chat = new ChatHistory();
        chat.AddSystemMessage("You are an expert analyst. Generate a professional summary of the key insights, data discoveries, and decisions from this conversation. Format as a Markdown report with sections: 'Executive Summary', 'Key Findings', 'Data Analysis', 'Next Steps'.");

        // Add context (last 20 turns or so)
        foreach (var msg in session.ContextWindow)
        {
            chat.Add(msg);
        }

        chat.AddUserMessage("Generate the insights report now.");

        var chatService = _kernel.GetRequiredService<IChatCompletionService>();
        var result = await chatService.GetChatMessageContentAsync(chat, kernel: _kernel);

        return result.Content ?? "Failed to generate summary.";
    }

    private static string EscapeCsv(string field)
    {
        if (string.IsNullOrEmpty(field)) return "";
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
        {
            return $"\"{field.Replace("\"", "\"\"")}\"";
        }
        return field;
    }
}
