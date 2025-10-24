# Testing Guide: User Story 1 MVP Validation

**Feature**: Local Excel Conversational Agent  
**User Story**: US1 - Basic Workbook Querying  
**Tasks**: T078-T084 (Validation & Integration)  
**Created**: October 23, 2025

---

## Overview

This guide provides step-by-step instructions for manually validating the User Story 1 MVP implementation. All implementation code is complete (T001-T077), and these tests verify the end-to-end functionality.

---

## Prerequisites

### 1. Sample Workbook

✅ **COMPLETED**: Sample workbook created at:
```
test-data/sample-workbook.xlsx
```

**Contents**:
- **Sales** sheet: 10 orders (SalesTable) with OrderID, Product, Region, Amount, Date
- **Inventory** sheet: 5 products (InventoryTable) with ProductID, Product, Stock, Warehouse
- **Returns** sheet: 3 returns (ReturnsTable) with ReturnID, Product, Reason, Date

### 2. Local LLM Server

**Option A - LM Studio** (Recommended):
1. Download from [lmstudio.ai](https://lmstudio.ai/)
2. Install and launch
3. Download model: `phi-4-mini-reasoning` (~7GB)
4. Start local server:
   - Click "Local Server" tab
   - Click "Start Server"
   - Default endpoint: `http://localhost:1234`
   - Note the model identifier (e.g., `phi-4-mini-reasoning`)

**Option B - Ollama**:
```powershell
# Install Ollama
winget install Ollama.Ollama

# Pull model
ollama pull phi4

# Start server (runs automatically on localhost:11434)
```

### 3. Application Configuration

Update `src/ExcelMcp.ChatWeb/appsettings.Development.json`:

```json
{
  "LlmStudio": {
    "BaseUrl": "http://localhost:1234",  // or http://localhost:11434 for Ollama
    "Model": "phi-4-mini-reasoning",     // or "phi4" for Ollama
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
  }
}
```

### 4. Build Verification

```powershell
# From repository root
cd 'd:\GitHub Projects\spec-kit-demos\local-workbook-mcp'

# Build solution
dotnet build ExcelLocalMcp.sln

# Expected: Build succeeded with 0 errors, 0 warnings
```

---

## Starting the Application

### Terminal 1: Run ChatWeb Application

```powershell
cd 'd:\GitHub Projects\spec-kit-demos\local-workbook-mcp'

# Run application
dotnet run --project src/ExcelMcp.ChatWeb

# Wait for: "Now listening on: http://localhost:5000"
# Or check configured port in output
```

### Browser: Open Chat Interface

1. Open browser (Chrome, Edge, Firefox)
2. Navigate to: `http://localhost:5000` or `http://localhost:5000/chat`
3. You should see:
   - Page title: "📊 Excel Conversational Agent"
   - Sidebar with "Workbook Selector"
   - Main area with message container
   - Input box for queries

---

## Test Scenarios (T078-T084)

### T078: Test Workbook Load Flow ✅

**Objective**: Verify file selection → metadata display → ready message

**Steps**:
1. Click "Choose File" button in sidebar
2. Navigate to: `d:\GitHub Projects\spec-kit-demos\local-workbook-mcp\test-data\`
3. Select `sample-workbook.xlsx`
4. Click "Load Workbook" button

**Expected Results**:
- ✓ Loading indicator appears
- ✓ Workbook info section displays:
  - Workbook name: "sample-workbook.xlsx"
  - Sheet count: 3
  - Sheet names: Sales, Inventory, Returns
- ✓ System message appears: "Workbook loaded successfully. Ready for queries."
- ✓ Suggested queries appear (e.g., "What sheets are in this workbook?")
- ✓ Query input is enabled

**Logs to Check** (`logs/agent-<date>.log`):
```json
{
  "@t": "...",
  "@mt": "Workbook loaded",
  "WorkbookPath": "...\\sample-workbook.xlsx",
  "SheetCount": 3
}
```

**Status**: ⬜ PASS / ❌ FAIL

**Notes**:
_____________________________________________

---

### T079: Test Basic Query ✅

**Objective**: Verify basic structure query works end-to-end

**Prerequisites**: Workbook loaded (T078 passed)

**Steps**:
1. In query input box, type: `What sheets are in this workbook?`
2. Click "Send" button or press Enter

**Expected Results**:
- ✓ User message appears in chat with timestamp
- ✓ Loading indicator shows "Processing query..."
- ✓ Assistant response appears within 10 seconds
- ✓ Response lists sheet names: "Sales, Inventory, Returns" (or similar natural language)
- ✓ Correlation ID displayed in message metadata

**Logs to Check**:
```json
{
  "@mt": "Query submitted",
  "Query": "What sheets are in this workbook?",
  "CorrelationId": "..."
}
{
  "@mt": "Tool invoked",
  "ToolName": "WorkbookStructurePlugin-ListWorkbookStructure",
  "Success": true
}
{
  "@mt": "Response generated",
  "ContentType": "Text",
  "ProcessingTimeMs": "<10000"
}
```

**Status**: ⬜ PASS / ❌ FAIL

**Notes**:
_____________________________________________

---

### T080: Test Data Retrieval ✅

**Objective**: Verify table data retrieval and HTML rendering

**Prerequisites**: Workbook loaded (T078 passed)

**Steps**:
1. Type query: `Show me the first 10 rows of the Sales table`
2. Click Send

**Expected Results**:
- ✓ Assistant response appears within 30 seconds
- ✓ Response contains HTML table with:
  - Table header: OrderID, Product, Region, Amount, Date
  - 10 data rows (all sales records from sample data)
  - Table metadata: Sheet name, row count
- ✓ Table styling applied (borders, hover effects, alternating row colors)
- ✓ ContentType: Table (visible in message metadata)

**Visual Validation**:
- Table should be readable and well-formatted
- No HTML escaping issues (no raw `<table>` tags visible)
- Data should match sample workbook contents

**Logs to Check**:
```json
{
  "@mt": "Tool invoked",
  "ToolName": "DataRetrievalPlugin-PreviewTable",
  "InputParameters": {"name": "SalesTable", "rowCount": 10}
}
{
  "@mt": "Response generated",
  "ContentType": "Table"
}
```

**Status**: ⬜ PASS / ❌ FAIL

**Notes**:
_____________________________________________

---

### T081: Test Error - No Workbook Loaded ✅

**Objective**: Verify error handling when no workbook is loaded

**Prerequisites**: Clear conversation or start fresh session

**Steps**:
1. **Without loading a workbook**, type query: `What sheets are in this workbook?`
2. Click Send

**Expected Results**:
- ✓ Error message appears (NOT exception stack trace)
- ✓ Error indicates no workbook loaded
- ✓ Suggested action: "Please load a workbook first"
- ✓ Error styling applied (red/error color scheme)
- ✓ Correlation ID present for troubleshooting

**Error Message Example**:
```
Error: No workbook is currently loaded. Please select and load an Excel file using the Workbook Selector.

Correlation ID: abc123-def456-...
```

**Logs to Check**:
```json
{
  "@mt": "Error occurred",
  "ErrorCode": "NO_WORKBOOK",
  "CorrelationId": "...",
  "SanitizedMessage": "..."
}
```

**Status**: ⬜ PASS / ❌ FAIL

**Notes**:
_____________________________________________

---

### T082: Test Error - Corrupted File ✅

**Objective**: Verify error handling for invalid Excel files

**Prerequisites**: None

**Steps**:
1. Create a corrupted file:
   ```powershell
   # Create text file with .xlsx extension
   "This is not an Excel file" | Out-File -FilePath ".\test-data\corrupted.xlsx"
   ```
2. In UI, load `corrupted.xlsx`
3. Observe error handling

**Expected Results**:
- ✓ Error message appears during load (NOT during query)
- ✓ Error is user-friendly: "Unable to load workbook. The file may be corrupted or not a valid Excel file."
- ✓ NO raw exception message or stack trace visible
- ✓ Correlation ID displayed
- ✓ Workbook remains in "not loaded" state

**Logs to Check**:
```json
{
  "@mt": "Workbook load failed",
  "ErrorCode": "WORKBOOK_LOAD_FAILED",
  "FilePath": "...\\corrupted.xlsx",
  "Exception": "..." // Should be logged but NOT shown to user
}
```

**Status**: ⬜ PASS / ❌ FAIL

**Notes**:
_____________________________________________

---

### T083: Test Query Timeout ⏱️

**Objective**: Verify 30-second timeout handling

**Prerequisites**: Workbook loaded

**Challenge**: This is difficult to test without simulating slow MCP responses

**Option A - Manual Simulation**:
1. Add artificial delay to MCP client code temporarily:
   ```csharp
   // In McpClientHost.cs or plugin
   await Task.Delay(35000); // Simulate 35-second delay
   ```
2. Rebuild and run
3. Submit query
4. Observe timeout behavior

**Option B - LLM Slow Response**:
1. Ask a very complex query that might take time:
   `Analyze all data across all sheets, calculate statistics for every column, and provide detailed insights`
2. Observe if timeout occurs

**Expected Results**:
- ✓ After 30 seconds, timeout error appears
- ✓ Error message indicates timeout occurred
- ✓ Suggested action: "Please try again with a simpler query or check your LLM server"
- ✓ User can retry without reloading workbook

**Logs to Check**:
```json
{
  "@mt": "Query timeout",
  "TimeoutSeconds": 30,
  "ErrorCode": "QUERY_TIMEOUT"
}
```

**Status**: ⬜ PASS / ❌ FAIL / ⚠️ SKIPPED

**Notes**:
_____________________________________________

---

### T084: End-to-End Validation (Full Quickstart) 📋

**Objective**: Validate all scenarios from quickstart.md work correctly

**Prerequisites**: Fresh application start, LLM server running

**Test Sequence**:

**1. Load Workbook**:
- Load `sample-workbook.xlsx`
- Verify metadata display

**2. Structure Queries**:
- "What sheets are in this workbook?"
- "What tables are available?"
- "Tell me about the Sales sheet"

**3. Data Retrieval Queries**:
- "Show me the first 5 rows of the Sales table"
- "What products are in the Inventory?"
- "Show me all returns"

**4. Search Queries**:
- "Search for 'Laptop' across all sheets"
- "Find all orders from the North region"

**5. Aggregation Queries**:
- "What's the total sales amount?"
- "How many products are in stock?"

**6. Multi-Turn Context**:
- "What's in the Sales sheet?"
- "Show me the top 3 rows" (should understand "Sales" from context)

**7. Conversation Management**:
- Conduct 5+ turn conversation
- Verify context maintained
- Click "Clear History"
- Verify history cleared but workbook still loaded

**Expected Results**:
- ✓ All queries return accurate responses
- ✓ Data matches workbook contents
- ✓ Tables render correctly
- ✓ Context maintained across turns
- ✓ Clear history works without losing workbook
- ✓ No crashes or exceptions
- ✓ Performance within acceptable limits (queries < 30s)

**Status**: ⬜ PASS / ❌ FAIL

**Notes**:
_____________________________________________

---

## Success Criteria Validation

After completing T078-T084, verify these success criteria from spec.md:

| ID | Criteria | Target | Status | Notes |
|----|----------|--------|--------|-------|
| SC-001 | Workbook load time | < 5 seconds | ⬜ | |
| SC-002 | Basic query response | < 10 seconds | ⬜ | |
| SC-003 | Structure query accuracy | 90% | ⬜ | |
| SC-004 | Data retrieval accuracy | 95% | ⬜ | |
| SC-005 | Complete workflow | < 3 minutes | ⬜ | |
| SC-007 | First query without help | 80% users | ⬜ | (Self-assessment) |
| SC-008 | Context maintained | 5+ queries | ⬜ | |
| SC-009 | Helpful error messages | 100% | ⬜ | |

---

## Troubleshooting

### LLM Not Responding

**Symptoms**: Loading indicator stuck, no response after 30+ seconds

**Checks**:
1. ✓ LM Studio is running and model loaded
2. ✓ Server started in LM Studio (green "Server Running" indicator)
3. ✓ `appsettings.Development.json` BaseUrl matches LM Studio port
4. ✓ Model name in config matches loaded model

**Test LLM directly**:
```powershell
curl http://localhost:1234/v1/models
# Should return model list
```

### MCP Server Issues

**Symptoms**: Tool invocation errors, "NO_WORKBOOK" errors

**Checks**:
1. ✓ Workbook file exists and is valid .xlsx
2. ✓ File path is accessible (no permission issues)
3. ✓ Check logs for MCP-related errors

### Blazor Circuit Disconnects

**Symptoms**: UI becomes unresponsive, no messages sent

**Checks**:
1. Browser console (F12) for errors
2. Look for WebSocket connection errors
3. Check server logs for exceptions
4. Refresh browser to reconnect

### Build Errors

**If build fails after pulling latest**:
```powershell
# Clean and rebuild
dotnet clean
dotnet build --no-incremental
```

---

## Reporting Results

### For Each Test (T078-T084)

Fill in:
- ✅ **Status**: PASS / FAIL / SKIPPED
- 📝 **Notes**: Any observations, issues, or deviations

### Overall Assessment

After completing all tests:

**MVP Readiness**:
- ⬜ **READY FOR RELEASE**: All tests passed, meets success criteria
- ⬜ **NEEDS WORK**: Some tests failed, requires fixes before release
- ⬜ **BLOCKED**: Critical issues prevent validation

**Known Issues**:
1. _____________________________________________
2. _____________________________________________
3. _____________________________________________

**Recommendations**:
- _____________________________________________
- _____________________________________________

---

## Next Steps

### If All Tests Pass ✅

1. Mark T078-T084 as [X] complete in tasks.md
2. Commit validation results and this testing guide
3. **DECISION POINT**:
   - **Option A**: Deploy/Demo MVP (User Story 1 is shippable!)
   - **Option B**: Proceed to Phase 4 (User Story 2 - Multi-Turn Context)

### If Tests Fail ❌

1. Document failures in detail (logs, screenshots)
2. Create bug reports with:
   - Test scenario (T0XX)
   - Expected vs actual results
   - Logs and correlation IDs
   - Reproduction steps
3. Fix critical issues before proceeding to next user story
4. Re-run validation after fixes

---

## Appendix: Sample Queries

### Basic Structure Queries
- "What sheets are in this workbook?"
- "List all sheets"
- "What tables are available?"
- "Tell me about the structure of this workbook"

### Data Retrieval Queries
- "Show me the first 10 rows of the Sales table"
- "Display all inventory data"
- "What's in the Returns sheet?"
- "Show me sales data"

### Search Queries
- "Find all mentions of Laptop"
- "Search for 'North' in the workbook"
- "Show me all defective returns"

### Aggregation Queries
- "What's the total sales amount?"
- "How many products are in stock?"
- "Count the number of returns"

### Multi-Turn Examples
1. "What's in the Sales sheet?"
2. "Show me the top 5 rows"
3. "What about the North region?"

---

**Validation Completed**: ________________ (Date)  
**Validated By**: ________________  
**MVP Status**: ⬜ PASS / ❌ FAIL
