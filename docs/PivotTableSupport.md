# Pivot Table Support

## Overview

The `excel-analyze-pivot` tool enables analysis of Excel pivot tables, extracting their structure, configuration, and aggregated data through the MCP server.

## Tool: excel-analyze-pivot

### Description
Analyze pivot tables in a worksheet, including structure, fields, and aggregated data.

### Input Schema

```json
{
  "worksheet": "string (required)",
  "pivotTable": "string (optional)",
  "includeFilters": "boolean (optional, default: true)",
  "maxRows": "integer (optional, default: 100, max: 1000)"
}
```

### Parameters

- **worksheet** (required): The worksheet name containing the pivot table(s)
- **pivotTable** (optional): Specific pivot table name to analyze. If omitted, all pivot tables in the worksheet are analyzed
- **includeFilters** (optional): Whether to include filter field information in the response (default: true)
- **maxRows** (optional): Maximum number of data rows to extract from the pivot table (default: 100, max: 1000)

### Response Structure

```json
{
  "pivotTables": [
    {
      "name": "string",
      "worksheetName": "string",
      "sourceWorksheet": "string",
      "sourceRange": "string",
      "rowFields": [
        {
          "name": "string",
          "sourceName": "string",
          "function": "string"
        }
      ],
      "columnFields": [...],
      "dataFields": [
        {
          "name": "string",
          "sourceName": "string",
          "function": "Sum|Average|Count|..."
        }
      ],
      "filterFields": [...],
      "data": [
        {
          "values": {
            "fieldName": "value"
          }
        }
      ]
    }
  ]
}
```

### Example Usage

#### CLI Client
```bash
dotnet run --project src/ExcelMcp.Client -- analyze-pivot SalesPivot --pivot "SalesSummary"
```

#### Semantic Kernel Agent
```
> What pivot tables exist in the Sales worksheet?
> Analyze the SalesSummary pivot table
> Show me the pivot table structure for regional sales
```

### Use Cases

1. **Pivot Discovery**: Find all pivot tables in a workbook
2. **Structure Analysis**: Understand how data is grouped and aggregated
3. **Data Extraction**: Get the summarized/aggregated values from pivot tables
4. **Cross-sheet Analysis**: Identify source data for pivot tables
5. **Filter Inspection**: See what filters are applied to pivot data

## Implementation Details

### New Contracts (ExcelMcp.Contracts)

- **PivotTableArguments**: Input parameters for pivot analysis
- **PivotTableResult**: Container for analyzed pivot tables
- **PivotTableInfo**: Detailed pivot table structure and data
- **PivotFieldInfo**: Field configuration (row/column/data/filter)
- **PivotDataRow**: Aggregated data rows from the pivot
- **PivotTableMetadata**: Lightweight metadata for pivot table discovery

### Service Method (ExcelWorkbookService)

```csharp
public async Task<PivotTableResult> AnalyzePivotTablesAsync(
    PivotTableArguments arguments,
    CancellationToken cancellationToken)
```

### MCP Tool Registration

The tool is registered in `McpServer` as `excel-analyze-pivot` with full JSON schema for input validation.

## Metadata Integration

Pivot tables now appear in the workbook metadata returned by `excel-list-structure`:

```json
{
  "worksheets": [
    {
      "name": "Sales",
      "tables": [...],
      "pivotTables": [
        {
          "name": "SalesSummary",
          "worksheetName": "Sales",
          "sourceRange": "...",
          "rowFieldCount": 2,
          "columnFieldCount": 1,
          "dataFieldCount": 1
        }
      ]
    }
  ]
}
```

## Testing

### Test Workbook

A test workbook with pivot table data has been created at `test-data/SalesWithPivot.xlsx` containing:

- **SalesData** worksheet: Source data table with Region, Product, SalesPerson, Amount, Quarter
- **SalesPivot** worksheet: Pivot table analyzing sales by Region/Product × Quarter

### Test Queries

```
> load test-data/SalesWithPivot.xlsx
> what worksheets are in this workbook?
> analyze the pivot table in SalesPivot
> what are the total sales by region?
```

## Limitations

Due to ClosedXML API constraints:
- Source range information is limited
- Some pivot table features may not be fully accessible
- Pivot table refresh is not supported

## Future Enhancements

See [FutureFeatures.md](FutureFeatures.md) for planned improvements:
- Pivot table creation/modification
- Drill-down to source data
- Pivot cache analysis
- Calculated fields/items support
- Slicer and filter integration
