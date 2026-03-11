# ChatWeb Manual Test Checklist
<!-- scripts/tests/manual-test-chatweb.md -->

This checklist covers every major user interaction in the Blazor ChatWeb app.
Work through each section in order; **check the box** when a step passes.

---

## Prerequisites

- [ ] Solution built: `dotnet build`
- [ ] Test data exists in `test-data/`  
      Generate if missing: `pwsh scripts/create-sample-workbooks.ps1`
- [ ] Ollama is running (`ollama serve`) or LM Studio is running on port 1234
- [ ] `appsettings.Development.json` points to your LLM (BaseUrl / Model)
- [ ] `ExcelMcp.ServerPath` in appsettings points to the built Server binary  
      Quick check: `ls src/ExcelMcp.Server/bin/Debug/net10.0/ExcelMcp.Server`
- [ ] App started: `./run-chatweb.sh`  
      Browser open to: **http://localhost:5000**

---

## Section 1 — App Load & Workbook Selector

| # | Step | Expected | ✓ |
|---|------|----------|---|
| CW-01 | Open http://localhost:5000 | Chat page loads; sidebar shows **workbook selector**; chat area shows welcome message | ☐ |
| CW-02 | Click or browse to `test-data/ProjectTracking.xlsx` and confirm | Sidebar shows **"Loaded Workbook"** panel with filename; lists sheet names (Tasks, Projects, TimeLog) | ☐ |
| CW-03 | Confirm **"Suggested Questions"** appear in the sidebar | At least 1–3 query suggestion buttons are shown | ☐ |
| CW-04 | Scroll the sheet list — confirm 3 sheets displayed | Tasks, Projects, TimeLog each appear | ☐ |

---

## Section 2 — ProjectTracking.xlsx Conversations

Workbook loaded: `test-data/ProjectTracking.xlsx`

### 2a — Structure Discovery

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-05 | `What sheets are in this workbook?` | All three sheet names (Tasks, Projects, TimeLog) | ☐ |
| CW-06 | `Show me the structure of the workbook` | Table/worksheet metadata, column names for each table | ☐ |
| CW-07 | `What columns does the TasksTable have?` | TaskID, TaskName, Owner, Status, Priority, DueDate | ☐ |

### 2b — Read / Search

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-08 | `What tasks are high priority?` | 5 high-priority tasks (Setup Database, Design UI Mockups, Security Audit, Deploy to Staging, Fix Database Backup) | ☐ |
| CW-09 | `Show me all tasks assigned to Alice` | Alice's 4 tasks (Setup Database, Write Documentation, Deploy to Staging, Update Dependencies) | ☐ |
| CW-10 | `Show me tasks that are In Progress` | 3 tasks (Design UI Mockups, Write Documentation, Fix Database Backup) | ☐ |
| CW-11 | `Which project has the highest budget?` | Mobile App — $75,000 | ☐ |
| CW-12 | `How many hours did Bob log?` | 20 hours (3 entries: 5+7+8) | ☐ |
| CW-13 | Click a **Suggested Question** button | The question populates the input field and sends, producing a relevant response | ☐ |

### 2c — Multi-turn Context

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-14 | `Show me all high-priority tasks` | High-priority results | ☐ |
| CW-15 | (follow-up) `Which of those are not started?` | Filters the previous results; Not Started + High priority tasks | ☐ |
| CW-16 | Check sidebar **context turns counter** | Counter increments correctly (e.g. "2 / 5 turns") | ☐ |

---

## Section 3 — EmployeeDirectory.xlsx Conversations

Switch workbook to `test-data/EmployeeDirectory.xlsx` (use the workbook selector).

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-17 | (after loading) `What sheets are in this workbook?` | Employees, Departments | ☐ |
| CW-18 | `Who works in Engineering?` | Alice Johnson, Bob Smith, Charlie Brown, Henry Wilson | ☐ |
| CW-19 | `What is the average salary by department?` | Salary figures per department (Engineering highest avg) | ☐ |
| CW-20 | `Who are the managers?` | Diana Prince (Sales), Frank Miller (HR), and possibly Bob Smith / Grace Lee | ☐ |
| CW-21 | `Which department has the largest budget?` | Engineering — $500,000 | ☐ |
| CW-22 | `Who was hired most recently?` | Henry Wilson (2024) | ☐ |

---

## Section 4 — BudgetTracker.xlsx Conversations

Switch workbook to `test-data/BudgetTracker.xlsx`.

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-23 | `What is the total income?` | $31,500 | ☐ |
| CW-24 | `Show me all software expenses` | Cloud Provider entries ($1,200 × 2) | ☐ |
| CW-25 | `What's our biggest expense?` | Office Rent — $3,000 | ☐ |
| CW-26 | `Which client generated the most revenue?` | Client B — $15,500 | ☐ |
| CW-27 | `What is the profit (income minus expenses)?` | ~$20,965 ($31,500 − $10,535) | ☐ |

---

## Section 5 — SalesWithPivot.xlsx (Pivot Feature)

Switch workbook to `test-data/SalesWithPivot.xlsx`  
_(requires `pwsh scripts/create-pivot-test-workbook.ps1`)_

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-28 | `What sheets are in this workbook?` | SalesData, SalesPivot | ☐ |
| CW-29 | `Analyze the pivot table in the SalesPivot sheet` | SalesSummary pivot details — row fields, column fields, value field (Amount Sum) | ☐ |
| CW-30 | `What were John's total sales?` | $7,200 across 4 entries | ☐ |
| CW-31 | `How many sales were in Q4?` | 1 sale — John, West, Widget, $2,300 | ☐ |

---

## Section 6 — Write-Back Operations

> ⚠️ **Use the backup copies** for write tests, or be aware the workbook will be modified.  
> A timestamped backup is created automatically before each write.

Reload `test-data/ProjectTracking.xlsx`.

| # | Prompt to type | Expected response contains | ✓ |
|---|----------------|----------------------------|---|
| CW-32 | `Update cell G1 in the Tasks sheet to "Notes"` | Confirmation that G1 was updated; backup path mentioned | ☐ |
| CW-33 | `Change the budget for the Mobile App project to 80000` | Confirmation that the cell was updated in the Projects sheet | ☐ |
| CW-34 | `Create a new worksheet called Summary` | Confirmation that the Summary sheet was added | ☐ |
| CW-35 | Reload the workbook in the selector | New `Summary` sheet now appears in the sidebar sheet list | ☐ |

---

## Section 7 — Session Controls

| # | Step | Expected | ✓ |
|---|------|----------|---|
| CW-36 | Click **Summarize** button in the sidebar | AI produces a brief text summary of the conversation so far | ☐ |
| CW-37 | Click **Export** button | Browser downloads (or prompts to save) a `.txt` / `.md` conversation export | ☐ |
| CW-38 | Click **Clear History** | Conversation messages cleared; welcome message re-appears; context counter resets | ☐ |
| CW-39 | After clearing, send a new message | Conversation works normally from a clean state | ☐ |

---

## Section 8 — Error & Edge Cases

| # | Scenario | Expected | ✓ |
|---|----------|----------|---|
| CW-40 | Type a message before loading a workbook | Friendly error or prompt to load a workbook first | ☐ |
| CW-41 | Ask about a sheet that doesn't exist: `Show me the Inventory sheet` | Graceful "not found" message — no crash | ☐ |
| CW-42 | Ask a completely off-topic question: `What is the capital of France?` | Answer is given (LLM responds) or politely redirected to workbook topics | ☐ |
| CW-43 | Stop the Ollama/LM Studio service; send a message | Graceful error displayed — no crash, no blank response | ☐ |

---

## Result Summary

| Section | Pass | Fail | Notes |
|---------|------|------|-------|
| 1 App Load | | | |
| 2 ProjectTracking | | | |
| 3 EmployeeDirectory | | | |
| 4 BudgetTracker | | | |
| 5 SalesWithPivot | | | |
| 6 Write-Back | | | |
| 7 Session Controls | | | |
| 8 Edge Cases | | | |
| **Total** | | | |

**Tester:**  
**Date:**  
**Ollama/LM Studio model used:**  
**OS / Browser:**  
**Build SHA / version:**  
