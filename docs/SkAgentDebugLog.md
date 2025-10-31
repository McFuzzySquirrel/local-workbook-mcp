# Debug Logging Feature

## What It Does

The agent now shows you **exactly what tools it's calling** (or if it's making things up!).

## Example Output

**When tools ARE being called:**
```
[sample-workbook.xlsx | local-model] > show me laptop sales

─── Debug Log ───
🔄 Sending request to LLM...
🔧 Tool Called: search(query='Laptop', worksheet='null', table='null')
✅ Found 3 matching rows
─────────────────

The workbook contains 3 laptop sales:
- Order 1001: $1200
- Order 1005: $1200
- Order 1008: $1200
```

**When LLM is making up answers:**
```
[sample-workbook.xlsx | local-model] > what's the total for laptops?

─── Debug Log ───
🔄 Sending request to LLM...
⚠️  No tools were called - LLM answered directly
─────────────────

The total for laptops is $3600.  <-- FABRICATED!
```

## Debug Log Icons

- 🔄 **Sending request to LLM** - Query is being sent
- 🔧 **Tool Called** - Shows which tool and parameters
- ✅ **Returned** - Shows how many results came back
- ⚠️ **No tools called** - LLM answered without using data (BAD!)
- ❌ **Error** - Something went wrong

## What To Look For

### ✅ GOOD BEHAVIOR
```
🔧 Tool Called: preview_table(worksheet='Sales', table='null', rows=20)
✅ Returned 10 rows from Sales
```
The LLM is using tools to get actual data!

### ⚠️ BAD BEHAVIOR
```
⚠️  No tools were called - LLM answered directly
```
The LLM is making up answers without checking the data!

## Tool Call Examples

### list_structure
```
🔧 Tool Called: list_structure
✅ Returned metadata for 3 worksheets
```

### search
```
🔧 Tool Called: search(query='Laptop', worksheet='Sales', table='null')
✅ Found 5 matching rows
```

### preview_table
```
🔧 Tool Called: preview_table(worksheet='Sales', table='SalesData', rows=50)
✅ Returned 50 rows from Sales/SalesData
```

### get_workbook_summary
```
🔧 Tool Called: get_workbook_summary
✅ Returned summary with 3 worksheets
```

## Troubleshooting With Debug Logs

### Problem: Wrong calculations

**Look for:**
```
⚠️  No tools were called - LLM answered directly
```

**Fix:** Your model doesn't support function calling well. Try:
- phi-4
- llama-3.1-8b-instruct  
- gpt-4

### Problem: "No results found" when data exists

**Look for:**
```
🔧 Tool Called: search(query='Product: Laptop', ...)
✅ Found 0 matching rows
```

**Issue:** Model is searching for "Product: Laptop" instead of just "Laptop"

**Fix:** Better model or modify the system prompt

### Problem: Only previewing 20 rows when you need all

**Look for:**
```
🔧 Tool Called: preview_table(worksheet='Sales', table='null', rows=20)
✅ Returned 20 rows from Sales
```

**Issue:** Model used default row count

**Fix:** Ask explicitly: "get ALL rows from Sales table" or "get 100 rows"

## How It Works

1. `ExcelAgent` has a `DebugLog` list
2. Each tool in `ExcelPlugin` logs when it's called
3. After LLM responds, debug log is displayed
4. If no tools were called, you get a warning

This helps you understand:
- **Is the model using tools?**
- **What parameters is it passing?**
- **How much data is it getting?**
- **Is it fabricating answers?**

## Disabling Debug Logs

If you want to turn off debug logging, edit `Program.cs` and comment out this section:

```csharp
// Show debug log if there are entries
if (agent.DebugLog.Count > 0)
{
    AnsiConsole.MarkupLine("[dim]─── Debug Log ───[/]");
    foreach (var logEntry in agent.DebugLog)
    {
        AnsiConsole.MarkupLine($"[dim]{logEntry.EscapeMarkup()}[/]");
    }
    AnsiConsole.MarkupLine("[dim]─────────────────[/]");
    AnsiConsole.WriteLine();
}
```

Now you can see exactly what your model is doing under the hood!
