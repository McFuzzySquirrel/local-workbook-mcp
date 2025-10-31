# Create Multiple Sample Workbooks for Testing
# This script creates several test Excel workbooks with different types of data

param(
    [string]$OutputDir = ".\test-data"
)

# Create test data directory if it doesn't exist
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# Check if ClosedXML assembly is available from built project
$assemblyPath = ".\src\ExcelMcp.Server\bin\Debug\net9.0\ClosedXML.dll"
if (-not (Test-Path $assemblyPath)) {
    Write-Error "ClosedXML assembly not found. Please build the solution first: dotnet build"
    exit 1
}

Add-Type -Path $assemblyPath

Write-Host "`nCreating sample workbooks..." -ForegroundColor Cyan
Write-Host "=" * 60

# ===== WORKBOOK 1: Project Tracking =====
Write-Host "`n1. Creating ProjectTracking.xlsx..." -ForegroundColor Yellow

$workbook = New-Object ClosedXML.Excel.XLWorkbook

# Tasks sheet
$tasksSheet = $workbook.Worksheets.Add("Tasks")
$tasksSheet.Cell(1, 1).Value = "TaskID"
$tasksSheet.Cell(1, 2).Value = "TaskName"
$tasksSheet.Cell(1, 3).Value = "Owner"
$tasksSheet.Cell(1, 4).Value = "Status"
$tasksSheet.Cell(1, 5).Value = "Priority"
$tasksSheet.Cell(1, 6).Value = "DueDate"

$tasksData = @(
    @(1, "Setup Database", "Alice", "Completed", "High", "2024-10-15"),
    @(2, "Design UI Mockups", "Bob", "In Progress", "High", "2024-11-05"),
    @(3, "Implement Login", "Charlie", "Not Started", "Medium", "2024-11-10"),
    @(4, "Write Documentation", "Alice", "In Progress", "Low", "2024-11-20"),
    @(5, "Security Audit", "Bob", "Not Started", "High", "2024-11-07"),
    @(6, "Performance Testing", "Charlie", "Not Started", "Medium", "2024-11-15"),
    @(7, "Deploy to Staging", "Alice", "Not Started", "High", "2024-11-12"),
    @(8, "User Acceptance Test", "Bob", "Not Started", "Medium", "2024-11-18"),
    @(9, "Fix Database Backup", "Charlie", "In Progress", "High", "2024-11-05"),
    @(10, "Update Dependencies", "Alice", "Completed", "Low", "2024-10-25")
)

$row = 2
foreach ($task in $tasksData) {
    for ($i = 0; $i -lt $task.Length; $i++) {
        $tasksSheet.Cell($row, $i + 1).Value = $task[$i]
    }
    $row++
}
$tasksRange = $tasksSheet.Range(1, 1, $row - 1, 6)
$tasksTable = $tasksRange.CreateTable("TasksTable")

# Projects sheet
$projectsSheet = $workbook.Worksheets.Add("Projects")
$projectsSheet.Cell(1, 1).Value = "ProjectID"
$projectsSheet.Cell(1, 2).Value = "ProjectName"
$projectsSheet.Cell(1, 3).Value = "Manager"
$projectsSheet.Cell(1, 4).Value = "Budget"
$projectsSheet.Cell(1, 5).Value = "StartDate"
$projectsSheet.Cell(1, 6).Value = "EndDate"

$projectsData = @(
    @(101, "Website Redesign", "Alice", 50000, "2024-10-01", "2024-12-31"),
    @(102, "Mobile App", "Bob", 75000, "2024-09-15", "2025-01-15"),
    @(103, "Data Migration", "Charlie", 30000, "2024-11-01", "2024-12-15")
)

$row = 2
foreach ($project in $projectsData) {
    for ($i = 0; $i -lt $project.Length; $i++) {
        $projectsSheet.Cell($row, $i + 1).Value = $project[$i]
    }
    $row++
}
$projectsRange = $projectsSheet.Range(1, 1, $row - 1, 6)
$projectsTable = $projectsRange.CreateTable("ProjectsTable")

# TimeLog sheet
$timeLogSheet = $workbook.Worksheets.Add("TimeLog")
$timeLogSheet.Cell(1, 1).Value = "EntryID"
$timeLogSheet.Cell(1, 2).Value = "TaskID"
$timeLogSheet.Cell(1, 3).Value = "Employee"
$timeLogSheet.Cell(1, 4).Value = "Hours"
$timeLogSheet.Cell(1, 5).Value = "Date"

$timeLogData = @(
    @(1, 1, "Alice", 8, "2024-10-10"),
    @(2, 1, "Alice", 6, "2024-10-11"),
    @(3, 2, "Bob", 5, "2024-10-28"),
    @(4, 2, "Bob", 7, "2024-10-29"),
    @(5, 4, "Alice", 4, "2024-10-30"),
    @(6, 9, "Charlie", 6, "2024-10-30"),
    @(7, 2, "Bob", 8, "2024-10-31")
)

$row = 2
foreach ($entry in $timeLogData) {
    for ($i = 0; $i -lt $entry.Length; $i++) {
        $timeLogSheet.Cell($row, $i + 1).Value = $entry[$i]
    }
    $row++
}
$timeLogRange = $timeLogSheet.Range(1, 1, $row - 1, 5)
$timeLogTable = $timeLogRange.CreateTable("TimeLogTable")

$outputPath = Join-Path $OutputDir "ProjectTracking.xlsx"
$workbook.SaveAs($outputPath)
$workbook.Dispose()
Write-Host "   ✓ ProjectTracking.xlsx created" -ForegroundColor Green

# ===== WORKBOOK 2: Employee Directory =====
Write-Host "`n2. Creating EmployeeDirectory.xlsx..." -ForegroundColor Yellow

$workbook = New-Object ClosedXML.Excel.XLWorkbook

# Employees sheet
$employeesSheet = $workbook.Worksheets.Add("Employees")
$employeesSheet.Cell(1, 1).Value = "EmployeeID"
$employeesSheet.Cell(1, 2).Value = "FullName"
$employeesSheet.Cell(1, 3).Value = "Department"
$employeesSheet.Cell(1, 4).Value = "Position"
$employeesSheet.Cell(1, 5).Value = "Salary"
$employeesSheet.Cell(1, 6).Value = "HireDate"

$employeesData = @(
    @(1001, "Alice Johnson", "Engineering", "Senior Developer", 95000, "2020-03-15"),
    @(1002, "Bob Smith", "Engineering", "Lead Developer", 110000, "2019-01-10"),
    @(1003, "Charlie Brown", "Engineering", "Developer", 75000, "2022-06-01"),
    @(1004, "Diana Prince", "Sales", "Sales Manager", 85000, "2021-04-20"),
    @(1005, "Eve Davis", "Sales", "Sales Rep", 55000, "2023-02-14"),
    @(1006, "Frank Miller", "HR", "HR Manager", 78000, "2020-08-01"),
    @(1007, "Grace Lee", "Marketing", "Marketing Director", 92000, "2019-11-05"),
    @(1008, "Henry Wilson", "Engineering", "Junior Developer", 62000, "2024-01-08")
)

$row = 2
foreach ($emp in $employeesData) {
    for ($i = 0; $i -lt $emp.Length; $i++) {
        $employeesSheet.Cell($row, $i + 1).Value = $emp[$i]
    }
    $row++
}
$employeesRange = $employeesSheet.Range(1, 1, $row - 1, 6)
$employeesTable = $employeesRange.CreateTable("EmployeesTable")

# Departments sheet
$departmentsSheet = $workbook.Worksheets.Add("Departments")
$departmentsSheet.Cell(1, 1).Value = "DeptID"
$departmentsSheet.Cell(1, 2).Value = "DeptName"
$departmentsSheet.Cell(1, 3).Value = "Manager"
$departmentsSheet.Cell(1, 4).Value = "Budget"

$departmentsData = @(
    @(10, "Engineering", "Bob Smith", 500000),
    @(20, "Sales", "Diana Prince", 300000),
    @(30, "HR", "Frank Miller", 150000),
    @(40, "Marketing", "Grace Lee", 250000)
)

$row = 2
foreach ($dept in $departmentsData) {
    for ($i = 0; $i -lt $dept.Length; $i++) {
        $departmentsSheet.Cell($row, $i + 1).Value = $dept[$i]
    }
    $row++
}
$departmentsRange = $departmentsSheet.Range(1, 1, $row - 1, 4)
$departmentsTable = $departmentsRange.CreateTable("DepartmentsTable")

$outputPath = Join-Path $OutputDir "EmployeeDirectory.xlsx"
$workbook.SaveAs($outputPath)
$workbook.Dispose()
Write-Host "   ✓ EmployeeDirectory.xlsx created" -ForegroundColor Green

# ===== WORKBOOK 3: Budget Tracker =====
Write-Host "`n3. Creating BudgetTracker.xlsx..." -ForegroundColor Yellow

$workbook = New-Object ClosedXML.Excel.XLWorkbook

# Income sheet
$incomeSheet = $workbook.Worksheets.Add("Income")
$incomeSheet.Cell(1, 1).Value = "Date"
$incomeSheet.Cell(1, 2).Value = "Source"
$incomeSheet.Cell(1, 3).Value = "Category"
$incomeSheet.Cell(1, 4).Value = "Amount"

$incomeData = @(
    @("2024-10-01", "Client A", "Consulting", 5000),
    @("2024-10-05", "Client B", "Development", 8000),
    @("2024-10-10", "Client C", "Support", 3500),
    @("2024-10-15", "Client A", "Consulting", 5000),
    @("2024-10-20", "Client D", "Training", 2500),
    @("2024-10-25", "Client B", "Development", 7500)
)

$row = 2
foreach ($income in $incomeData) {
    for ($i = 0; $i -lt $income.Length; $i++) {
        $incomeSheet.Cell($row, $i + 1).Value = $income[$i]
    }
    $row++
}
$incomeRange = $incomeSheet.Range(1, 1, $row - 1, 4)
$incomeTable = $incomeRange.CreateTable("IncomeTable")

# Expenses sheet
$expensesSheet = $workbook.Worksheets.Add("Expenses")
$expensesSheet.Cell(1, 1).Value = "Date"
$expensesSheet.Cell(1, 2).Value = "Vendor"
$expensesSheet.Cell(1, 3).Value = "Category"
$expensesSheet.Cell(1, 4).Value = "Amount"

$expensesData = @(
    @("2024-10-02", "Office Supplies Co", "Supplies", 450),
    @("2024-10-03", "Cloud Provider", "Software", 1200),
    @("2024-10-07", "Coffee Shop", "Food", 85),
    @("2024-10-10", "Tech Vendor", "Equipment", 2500),
    @("2024-10-15", "Utilities Inc", "Utilities", 300),
    @("2024-10-18", "Cloud Provider", "Software", 1200),
    @("2024-10-22", "Marketing Agency", "Marketing", 1800),
    @("2024-10-28", "Office Rent", "Rent", 3000)
)

$row = 2
foreach ($expense in $expensesData) {
    for ($i = 0; $i -lt $expense.Length; $i++) {
        $expensesSheet.Cell($row, $i + 1).Value = $expense[$i]
    }
    $row++
}
$expensesRange = $expensesSheet.Range(1, 1, $row - 1, 4)
$expensesTable = $expensesRange.CreateTable("ExpensesTable")

$outputPath = Join-Path $OutputDir "BudgetTracker.xlsx"
$workbook.SaveAs($outputPath)
$workbook.Dispose()
Write-Host "   ✓ BudgetTracker.xlsx created" -ForegroundColor Green

# ===== Summary =====
Write-Host "`n" + ("=" * 60)
Write-Host "`n✓ All sample workbooks created in: $OutputDir" -ForegroundColor Green
Write-Host "`nWorkbooks created:" -ForegroundColor Cyan
Write-Host "  1. ProjectTracking.xlsx   - Tasks, projects, and time logs"
Write-Host "  2. EmployeeDirectory.xlsx - Employee data and departments"
Write-Host "  3. BudgetTracker.xlsx     - Income and expense tracking"
Write-Host "`nTry them with:" -ForegroundColor Cyan
Write-Host "  dotnet run --project src/ExcelMcp.SkAgent -- --workbook test-data/ProjectTracking.xlsx"
Write-Host "  Or use 'load' command in the CLI to switch between them!"
Write-Host ""
