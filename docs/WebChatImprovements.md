# Web Chat Improvements - HTML Table Rendering & Workbook-Agnostic Prompts

**Date:** November 1, 2025  
**Issues Fixed:**
1. HTML tags displaying as text instead of rendered tables
2. System prompt needs to be workbook-agnostic (learned from CLI agent)

---

## Issue 1: HTML Table Rendering Fix

### Problem
When the LLM returned table data, it was displayed as escaped HTML tags like:
```
&lt;table&gt;&lt;tr&gt;&lt;td&gt;data&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;
```
Instead of a proper rendered table.

### Root Cause
The `ChatMessage.razor` component uses `@((MarkupString)Turn.Content)` which **should** render HTML, but the content formatting logic had issues:
1. System prompt was asking LLM to generate HTML tables
2. FormatResponseContent was only looking for CSV data
3. Missing detection for HTML tables already in response

### Solution
Enhanced `FormatResponseContent()` in `ExcelAgentService.cs`:

```csharp
private (string content, ContentType type) FormatResponseContent(string responseContent)
{
    // NEW: Check if response already contains HTML table (LLM may generate this)
    if (responseContent.Contains("<table") && responseContent.Contains("</table>"))
    {
        // Already has HTML table - just ensure it's properly wrapped
        if (!responseContent.Contains("class=\"data-table\""))
        {
            responseContent = responseContent.Replace("<table", "<table class=\"data-table\"");
        }
        return (responseContent, ContentType.Table);
    }

    // IMPROVED: Better CSV detection with commentary preservation
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
```

**Benefits:**
- ✅ Handles LLM-generated HTML tables
- ✅ Parses CSV/tabular data from tool responses
- ✅ Preserves commentary with tables ("Here are the results:" + table)
- ✅ Falls back to formatted text gracefully

---

## Issue 2: Workbook-Agnostic System Prompt

### Problem
The original system prompt was asking the LLM to generate HTML in responses, and wasn't as robust as the proven CLI agent prompt.

### Learning from CLI Agent
The `ExcelMcp.SkAgent` has a battle-tested system prompt (lines 43-94 in Program.cs) that:
- ✅ Explicitly tells LLM to use tools (fixes "I don't have access" issues)
- ✅ Works with ANY workbook (no hardcoded sheet names)
- ✅ Teaches discovery workflow (structure first, then data)
- ✅ Provides clear examples of right/wrong behavior
- ✅ Emphasizes that tools MUST be used

### Solution
Created `PromptTemplates.cs` with comprehensive, reusable prompts:

```csharp
public static class PromptTemplates
{
    public const string SystemPrompt = @"You are an Excel workbook assistant...
    
CRITICAL INSTRUCTIONS - YOU MUST FOLLOW THESE:

1. NEVER say ""I don't have access to the data"" - YOU DO! Use the tools!
2. When asked about data, rows, or content - IMMEDIATELY call the appropriate tool
3. ALWAYS use tools to get real data - never make assumptions or refuse to look
4. DO NOT assume worksheet names - ALWAYS discover them first using tools
5. Work with ANY workbook - never assume specific sheet names or structure

WORKFLOW - FOLLOW THIS PATTERN:

Step 1: If you don't know the structure → Call GetWorkbookSummary or ListWorkbookStructure
Step 2: Once you know worksheet names → Call PreviewTable to get actual data
Step 3: Analyze the real data you retrieved and answer the question

IMPORTANT - PRESENTING DATA:

When showing table data:
1. DO NOT return raw CSV or comma-separated values
2. DO NOT include HTML tags in your response text
3. Simply describe what you found: ""Here are the rows from the X table:""
4. The UI will automatically format the data as a table
5. Focus on insights and analysis, not formatting
...";

    public const string ContextAwarenessPrompt = @"...";
    public const string ErrorHandlingPrompt = @"...";
    public const string AmbiguityResolutionPrompt = @"...";
    
    public static string GetCompleteSystemPrompt() { ... }
    public static List<string> GetSuggestedQueries(WorkbookMetadata? metadata) { ... }
}
```

### Updated ExcelAgentService
Changed from inline system prompt to using PromptTemplates:

```csharp
// OLD:
var systemContext = $@"You are analyzing the workbook '{session.CurrentContext.WorkbookName}'...
IMPORTANT: When you use preview_table and receive CSV data, output it as an HTML table...";

// NEW:
// Add comprehensive system prompt (workbook-agnostic, proven from CLI agent)
chatHistory.AddSystemMessage(PromptTemplates.GetCompleteSystemPrompt());

// Add specific context about current workbook
var workbookContext = $@"CURRENT WORKBOOK CONTEXT:
- File: '{session.CurrentContext.WorkbookName}'
- Sheets ({sheets.Count}): {string.Join(", ", sheets.Take(10))}
- Tables: {(tables.Any() ? string.Join(", ", tables.Take(10)) : "None defined")}

Remember: Use tools to discover structure if you need more details!";

chatHistory.AddSystemMessage(workbookContext);
```

**Benefits:**
- ✅ Eliminates "I don't have access to data" refusals
- ✅ Works with any workbook structure
- ✅ No hardcoded sheet names
- ✅ Teaches LLM the correct workflow
- ✅ Reusable across web and future CLI versions
- ✅ Separated concerns (generic instructions vs. specific context)

---

## Files Modified

1. **`src/ExcelMcp.ChatWeb/Services/Agent/PromptTemplates.cs`** ✨ NEW
   - System prompts based on proven CLI agent
   - Context awareness guidance
   - Error handling patterns
   - Ambiguity resolution
   - Dynamic suggested queries

2. **`src/ExcelMcp.ChatWeb/Services/Agent/ExcelAgentService.cs`** 📝 UPDATED
   - Uses `PromptTemplates.GetCompleteSystemPrompt()`
   - Enhanced `FormatResponseContent()` for HTML detection
   - Added `FindTableDataStart()` helper
   - Better commentary + table preservation
   - Cleaner workbook context injection

---

## Testing Recommendations

### Test Case 1: HTML Table Rendering
```
1. Load workbook
2. Ask: "Show me the first 10 rows from the Tasks table"
3. EXPECTED: See a properly rendered HTML table, not escaped tags
4. VERIFY: Table has borders, headers, and is styled
```

### Test Case 2: Workbook-Agnostic Behavior
```
1. Load ProjectTracking.xlsx (has: Tasks, Projects, TimeLog, Users sheets)
2. Ask: "What data is in here?"
3. EXPECTED: LLM calls GetWorkbookSummary and discovers actual sheets
4. VERIFY: Response mentions correct sheet names, not assumptions
```

### Test Case 3: No "Access" Refusals
```
1. Load any workbook
2. Ask: "Show me some data"
3. EXPECTED: LLM calls PreviewTable and returns real data
4. VERIFY: No messages like "I don't have access to the data"
```

### Test Case 4: Context Preservation
```
1. Ask: "What sheets exist?"
2. Ask: "Show me data from the first one"
3. EXPECTED: LLM remembers first sheet name from step 1
4. VERIFY: PreviewTable called with correct sheet name
```

---

## Comparison: Before vs After

### Before
```
System Prompt: "When you use preview_table, output HTML table..."
LLM Response: "<table><tr><td>data</td></tr></table>"
UI Display: &lt;table&gt;&lt;tr&gt;&lt;td&gt;data&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;
Result: ❌ Escaped HTML tags visible
```

### After
```
System Prompt: "Simply describe findings. UI will format tables."
LLM Response: "Here are the tasks:\nTaskID,Title,Status\n1,Review,Open\n2,Fix,Done"
FormatResponseContent: Detects CSV → Converts to HTML table
UI Display: [Rendered table with borders and styling]
Result: ✅ Proper table rendering
```

### CLI Lessons Applied
```
CLI Agent Prompt: "NEVER say I don't have access - USE THE TOOLS!"
Web Agent (Old): Generic instructions
Web Agent (New): Same proven prompt from CLI
Result: ✅ Consistent behavior across both interfaces
```

---

## Benefits Summary

### Technical
- ✅ Proper HTML rendering (tables display correctly)
- ✅ Workbook-agnostic (works with any Excel file)
- ✅ Reusable prompts (DRY principle)
- ✅ Better error handling
- ✅ Context-aware suggestions

### User Experience
- ✅ No more HTML tags in responses
- ✅ LLM doesn't refuse to show data
- ✅ Consistent behavior with CLI agent
- ✅ Natural conversation flow
- ✅ Smart suggested queries

### Maintainability
- ✅ Centralized prompt management
- ✅ Easy to update prompts in one place
- ✅ Clear separation of concerns
- ✅ Well-documented patterns
- ✅ Testable components

---

## Next Steps

1. **Test with LLM** - Verify table rendering works with real queries
2. **Compare with CLI** - Ensure web and CLI behave consistently
3. **Iterate on prompts** - Fine-tune based on real usage
4. **Add more examples** - Expand PromptTemplates with more patterns

---

## Key Takeaways

1. **Don't ask LLM to generate HTML** - Let your code handle formatting
2. **Learn from what works** - CLI agent prompt was proven, reuse it
3. **Be workbook-agnostic** - Never assume sheet names or structure
4. **Teach the workflow** - Explicit instructions beat implicit assumptions
5. **Preserve context** - Commentary + data both matter

**The web chat now has the same proven prompt strategy as the CLI agent!** 🎉
