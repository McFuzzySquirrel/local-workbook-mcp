# Sample Workbooks Reference

Three sample Excel workbooks have been created in `test-data/` for testing the agent.

## 1. ProjectTracking.xlsx

**Use Case:** Software development project management

### Tables:
- **TasksTable** (10 rows)
  - TaskID, TaskName, Owner, Status, Priority, DueDate
  
- **ProjectsTable** (3 rows)
  - ProjectID, ProjectName, Manager, Budget, StartDate, EndDate
  
- **TimeLogTable** (7 rows)
  - EntryID, TaskID, Employee, Hours, Date

### Example Questions:
```
> what tasks are high priority?
> show me all tasks assigned to Alice
> which tasks are overdue?
> how many hours did Bob log?
> what's the total budget across all projects?
> show me tasks that are "In Progress"
> which project has the highest budget?
> what tasks are due this week?
```

### Sample Data Highlights:
- 3 projects (Website Redesign, Mobile App, Data Migration)
- Tasks with various statuses (Completed, In Progress, Not Started)
- Priorities (High, Medium, Low)
- Time entries from multiple employees

---

## 2. EmployeeDirectory.xlsx

**Use Case:** HR and employee management

### Tables:
- **EmployeesTable** (8 rows)
  - EmployeeID, FullName, Department, Position, Salary, HireDate
  
- **DepartmentsTable** (4 rows)
  - DeptID, DeptName, Manager, Budget

### Example Questions:
```
> who works in Engineering?
> show me all employees hired in 2020
> what's the average salary by department?
> who are the managers?
> how many people work in each department?
> show me the highest paid employees
> which department has the largest budget?
> who was hired most recently?
```

### Sample Data Highlights:
- 4 departments (Engineering, Sales, HR, Marketing)
- Salary range from $55k to $110k
- Various positions (Developer, Manager, Director, etc.)
- Hire dates from 2019 to 2024

---

## 3. BudgetTracker.xlsx

**Use Case:** Financial tracking and budgeting

### Tables:
- **IncomeTable** (6 rows)
  - Date, Source, Category, Amount
  
- **ExpensesTable** (8 rows)
  - Date, Vendor, Category, Amount

### Example Questions:
```
> what's the total income for October?
> show me all software expenses
> what's our biggest expense?
> which client generated the most revenue?
> what's the profit (income minus expenses)?
> show me all transactions from Client B
> what categories of expenses do we have?
> how much did we spend on marketing?
```

### Sample Data Highlights:
- Income from 4 clients
- Categories: Consulting, Development, Support, Training
- Expense categories: Supplies, Software, Equipment, Rent, etc.
- Total income: $31,500
- Total expenses: $10,535

---

## Quick Start

### Create the workbooks:
```pwsh
pwsh -File scripts/create-sample-workbooks.ps1
```

### Test with the CLI agent:
```pwsh
# Start with ProjectTracking
dotnet run --project src/ExcelMcp.SkAgent -- --workbook test-data/ProjectTracking.xlsx

# Inside the agent, switch between workbooks:
> load test-data/EmployeeDirectory.xlsx
> load test-data/BudgetTracker.xlsx
```

### Test workbook switching:
```
[ProjectTracking.xlsx | gpt-4] > what tasks are high priority?
... (shows tasks) ...

[ProjectTracking.xlsx | gpt-4] > load test-data/EmployeeDirectory.xlsx
✓ Loaded workbook: EmployeeDirectory.xlsx

[EmployeeDirectory.xlsx | gpt-4] > who works in Engineering?
... (shows employees) ...

[EmployeeDirectory.xlsx | gpt-4] > switch
Enter path to new workbook: test-data/BudgetTracker.xlsx
✓ Loaded workbook: BudgetTracker.xlsx

[BudgetTracker.xlsx | gpt-4] > what's the total income?
... (calculates income) ...
```

## Testing Recommendations

### Test Basic Queries:
1. **List structure**: "what tables exist?" 
2. **Search**: "show me all laptop sales" (in original sample-workbook.xlsx)
3. **Preview**: "show me the first 5 rows of TasksTable"
4. **Summary**: "summarize this workbook"

### Test Calculations:
1. "what's the total budget for all projects?"
2. "calculate total expenses in October"
3. "what's the average salary in Engineering?"
4. "how many hours were logged total?"

### Test Filtering:
1. "show me only high priority tasks"
2. "which employees earn over $80,000?"
3. "what expenses are categorized as Software?"

### Test Multi-Table Analysis:
1. "which project has the most time logged?"
2. "how many employees are in each department and what's their total salary?"
3. "compare income vs expenses by category"

### Watch for:
- ✅ **Tool calls in debug log** - Should show `preview_table` or `search`
- ⚠️ **Correct calculations** - Model should use actual data, not fabricate
- ✅ **Case-insensitive search** - "laptop" should find "Laptop"
- ✅ **Empty worksheet params** - Should search across all sheets

## Original Sample Workbook

The original `sample-workbook.xlsx` is still available with:
- Sales data (laptops, mice, keyboards, etc.)
- Inventory data
- Returns data

Use it for e-commerce/retail testing scenarios.

---

**Now you have 4 workbooks to test different scenarios!** 🎉
