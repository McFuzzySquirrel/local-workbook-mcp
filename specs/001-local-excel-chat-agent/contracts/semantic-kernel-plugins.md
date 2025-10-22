# Semantic Kernel Plugin Contracts

**Feature**: Local Excel Conversational Agent  
**Date**: October 22, 2025

## Overview

This document defines the Semantic Kernel plugin interfaces that wrap MCP tools for workbook analysis. Each plugin method corresponds to one or more MCP tool operations.

---

## WorkbookStructurePlugin

Provides workbook metadata and structural information.

### ListWorkbookStructure

**MCP Tool**: `excel-list-structure`

```csharp
[KernelFunction("list_workbook_structure")]
[Description("Lists all sheets, tables, and structural information from the loaded workbook")]
[return: Description("JSON string containing sheets, tables, row/column counts")]
public async Task<string> ListWorkbookStructure()
```

**Input**: None (operates on currently loaded workbook)

**Output** (JSON string):
```json
{
  "sheets": [
    {
      "name": "Sales",
      "rowCount": 1523,
      "columnCount": 12,
      "tables": ["SalesData", "Targets"]
    },
    {
      "name": "Inventory",
      "rowCount": 845,
      "columnCount": 8,
      "tables": ["Products"]
    }
  ],
  "workbookName": "Q4_Report.xlsx",
  "totalSheets": 2,
  "totalTables": 3
}
```

**Error Handling**:
- No workbook loaded → Return error JSON with code "NO_WORKBOOK"
- MCP call fails → Return error JSON with code "MCP_ERROR"

---

### GetSheetNames

**MCP Tool**: Derived from `excel-list-structure`

```csharp
[KernelFunction("get_sheet_names")]
[Description("Gets a simple list of all worksheet names in the workbook")]
[return: Description("Comma-separated list of sheet names")]
public async Task<string> GetSheetNames()
```

**Input**: None

**Output**: `"Sales, Inventory, Dashboard, Summary"`

**Error Handling**: Same as ListWorkbookStructure

---

### GetTableInfo

**MCP Tool**: Derived from `excel-list-structure`

```csharp
[KernelFunction("get_table_info")]
[Description("Gets information about tables in a specific worksheet")]
public async Task<string> GetTableInfo(
    [Description("The worksheet name to query")] string sheetName
)
```

**Input**:
- `sheetName`: Name of the worksheet (case-sensitive)

**Output** (JSON string):
```json
{
  "sheetName": "Sales",
  "tables": [
    {
      "name": "SalesData",
      "rowCount": 1500,
      "columnCount": 10,
      "columns": ["Date", "Product", "Quantity", "Revenue", ...]
    }
  ]
}
```

**Error Handling**:
- Sheet not found → Return error JSON with code "SHEET_NOT_FOUND"
- No tables in sheet → Return empty tables array

---

## WorkbookSearchPlugin

Searches workbook content for specific text or patterns.

### SearchWorkbook

**MCP Tool**: `excel-search`

```csharp
[KernelFunction("search_workbook")]
[Description("Searches all worksheets for cells containing the specified text")]
public async Task<string> SearchWorkbook(
    [Description("The text to search for (case-insensitive)")] string searchText,
    [Description("Maximum results to return (default 50)")] int maxResults = 50
)
```

**Input**:
- `searchText`: Text to find (supports partial matches)
- `maxResults`: Limit results (default 50, max 500)

**Output** (JSON string):
```json
{
  "searchText": "Project X",
  "totalMatches": 23,
  "results": [
    {
      "sheetName": "Sales",
      "cellReference": "A15",
      "value": "Project X - Phase 1",
      "row": 15,
      "column": 1
    },
    {
      "sheetName": "Budget",
      "cellReference": "C42",
      "value": "Project X allocation",
      "row": 42,
      "column": 3
    }
  ],
  "truncated": false
}
```

**Error Handling**:
- Empty search text → Return error JSON with code "INVALID_INPUT"
- Search timeout (>30s) → Return partial results with truncated=true

---

### SearchInSheet

**MCP Tool**: `excel-search` with sheet filter

```csharp
[KernelFunction("search_in_sheet")]
[Description("Searches a specific worksheet for cells containing the text")]
public async Task<string> SearchInSheet(
    [Description("The worksheet name")] string sheetName,
    [Description("The text to search for")] string searchText,
    [Description("Maximum results (default 50)")] int maxResults = 50
)
```

**Input**:
- `sheetName`: Target worksheet
- `searchText`: Text to find
- `maxResults`: Result limit

**Output**: Same format as SearchWorkbook, but results only from specified sheet

**Error Handling**:
- Sheet not found → Return error JSON with code "SHEET_NOT_FOUND"
- Otherwise same as SearchWorkbook

---

## DataRetrievalPlugin

Retrieves actual data from worksheets and tables.

### PreviewTable

**MCP Tool**: `excel-preview-table`

```csharp
[KernelFunction("preview_table")]
[Description("Retrieves a preview of rows from a named table or worksheet")]
public async Task<string> PreviewTable(
    [Description("The sheet or table name")] string name,
    [Description("Number of rows to retrieve (default 10, max 100)")] int rowCount = 10,
    [Description("Starting row number (0-based, default 0)")] int startRow = 0
)
```

**Input**:
- `name`: Table name or sheet name
- `rowCount`: How many rows to fetch
- `startRow`: Offset for pagination

**Output** (JSON string):
```json
{
  "name": "SalesData",
  "columns": ["Date", "Product", "Quantity", "Revenue", "Region"],
  "rows": [
    ["2024-01-15", "Widget A", "150", "$15,000", "West"],
    ["2024-01-16", "Widget B", "200", "$30,000", "East"],
    ...
  ],
  "totalRows": 1500,
  "startRow": 0,
  "returnedRows": 10,
  "hasMore": true
}
```

**Error Handling**:
- Table/sheet not found → Return error JSON with code "NOT_FOUND"
- Invalid pagination → Return error JSON with code "INVALID_RANGE"
- Timeout → Return partial results with hasMore=false

---

### GetRowsInRange

**MCP Tool**: `excel-preview-table` with cell range

```csharp
[KernelFunction("get_rows_in_range")]
[Description("Retrieves data from a specific cell range in a worksheet")]
public async Task<string> GetRowsInRange(
    [Description("The worksheet name")] string sheetName,
    [Description("Cell range (e.g., 'A1:E10')")] string cellRange
)
```

**Input**:
- `sheetName`: Target worksheet
- `cellRange`: Excel cell range notation (e.g., "A1:E10")

**Output**: Same format as PreviewTable, with rows from specified range

**Error Handling**:
- Invalid range format → Return error JSON with code "INVALID_RANGE"
- Range too large (>1000 cells) → Return error JSON with code "RANGE_TOO_LARGE"

---

### CalculateAggregation

**MCP Tool**: Combination of `excel-preview-table` + client-side calculation

```csharp
[KernelFunction("calculate_aggregation")]
[Description("Calculates sum, average, min, or max for a numeric column")]
public async Task<string> CalculateAggregation(
    [Description("Table or sheet name")] string name,
    [Description("Column name or index")] string column,
    [Description("Aggregation type: sum, avg, min, max, count")] string aggregationType
)
```

**Input**:
- `name`: Table or sheet name
- `column`: Column identifier (name or 0-based index)
- `aggregationType`: One of: sum, avg, min, max, count

**Output** (JSON string):
```json
{
  "name": "SalesData",
  "column": "Revenue",
  "aggregationType": "sum",
  "result": 1234567.89,
  "rowCount": 1500,
  "format": "$#,##0.00"
}
```

**Error Handling**:
- Non-numeric column → Return error JSON with code "NOT_NUMERIC"
- Column not found → Return error JSON with code "COLUMN_NOT_FOUND"
- Invalid aggregation type → Return error JSON with code "INVALID_AGGREGATION"

---

## Error Response Format

All plugin methods return errors in consistent JSON format:

```json
{
  "error": true,
  "errorCode": "SHEET_NOT_FOUND",
  "message": "The requested sheet could not be found",
  "correlationId": "123e4567-e89b-12d3-a456-426614174000",
  "timestamp": "2025-10-22T14:30:00Z",
  "canRetry": false,
  "suggestedAction": "Check available sheets using list_workbook_structure"
}
```

**Standard Error Codes**:
- `NO_WORKBOOK`: No workbook currently loaded
- `SHEET_NOT_FOUND`: Referenced sheet doesn't exist
- `TABLE_NOT_FOUND`: Referenced table doesn't exist
- `COLUMN_NOT_FOUND`: Referenced column doesn't exist
- `INVALID_INPUT`: Invalid parameter value
- `INVALID_RANGE`: Invalid cell range format or out of bounds
- `RANGE_TOO_LARGE`: Requested range exceeds size limits
- `NOT_NUMERIC`: Operation requires numeric data but column is text
- `TIMEOUT`: Operation exceeded time limit
- `MCP_ERROR`: Underlying MCP tool call failed
- `UNKNOWN_ERROR`: Unexpected error occurred

---

## Plugin Registration

```csharp
// In Program.cs or service configuration
builder.Services.AddSingleton<WorkbookStructurePlugin>();
builder.Services.AddSingleton<WorkbookSearchPlugin>();
builder.Services.AddSingleton<DataRetrievalPlugin>();

// Register with Semantic Kernel
var kernel = builder.Services.AddKernel()
    .AddOpenAIChatCompletion(
        modelId: llmOptions.Model,
        endpoint: new Uri(llmOptions.BaseUrl),
        apiKey: "not-needed-for-local"
    )
    .Plugins.AddFromObject<WorkbookStructurePlugin>()
    .Plugins.AddFromObject<WorkbookSearchPlugin>()
    .Plugins.AddFromObject<DataRetrievalPlugin>()
    .Build();
```

---

## Usage Example (from LLM perspective)

User query: "What sheets are in the workbook?"

SK generates tool call:
```json
{
  "function": "list_workbook_structure",
  "parameters": {}
}
```

Plugin returns:
```json
{
  "sheets": [
    {"name": "Sales", "rowCount": 1523, "columnCount": 12, "tables": ["SalesData"]},
    {"name": "Inventory", "rowCount": 845, "columnCount": 8, "tables": ["Products"]}
  ],
  "workbookName": "Q4_Report.xlsx",
  "totalSheets": 2,
  "totalTables": 2
}
```

SK generates natural language response:
"The workbook Q4_Report.xlsx contains 2 sheets: Sales (with 1,523 rows and table SalesData) and Inventory (with 845 rows and table Products)."

---

## Testing Contracts

Each plugin method must have unit tests covering:

1. **Happy path**: Valid input → Expected output
2. **Error cases**: Invalid input → Proper error JSON
3. **Boundary conditions**: Max limits, empty results
4. **MCP failure**: Simulated MCP error → Proper handling

Example test structure:

```csharp
public class WorkbookStructurePluginTests
{
    [Fact]
    public async Task ListWorkbookStructure_WhenWorkbookLoaded_ReturnsStructure()
    {
        // Arrange
        var mockMcpClient = CreateMockClientWithWorkbook();
        var plugin = new WorkbookStructurePlugin(mockMcpClient);
        
        // Act
        var result = await plugin.ListWorkbookStructure();
        
        // Assert
        var data = JsonSerializer.Deserialize<WorkbookStructureResponse>(result);
        Assert.NotNull(data);
        Assert.Equal(2, data.TotalSheets);
    }
    
    [Fact]
    public async Task ListWorkbookStructure_WhenNoWorkbook_ReturnsError()
    {
        // Arrange
        var mockMcpClient = CreateMockClientWithNoWorkbook();
        var plugin = new WorkbookStructurePlugin(mockMcpClient);
        
        // Act
        var result = await plugin.ListWorkbookStructure();
        
        // Assert
        var error = JsonSerializer.Deserialize<ErrorResponse>(result);
        Assert.True(error.Error);
        Assert.Equal("NO_WORKBOOK", error.ErrorCode);
    }
}
```
