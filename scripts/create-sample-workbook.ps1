# Create Sample Workbook for Testing
# This script creates a test Excel workbook with sample data

param(
    [string]$OutputPath = ".\test-data\sample-workbook.xlsx"
)

# Create test data directory if it doesn't exist
$testDataDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $testDataDir)) {
    New-Item -ItemType Directory -Path $testDataDir -Force | Out-Null
}

# Check if ClosedXML assembly is available from built project
$assemblyPath = ".\src\ExcelMcp.Server\bin\Debug\net9.0\ClosedXML.dll"
if (-not (Test-Path $assemblyPath)) {
    Write-Error "ClosedXML assembly not found. Please build the solution first: dotnet build"
    exit 1
}

Add-Type -Path $assemblyPath

# Create workbook
$workbook = New-Object ClosedXML.Excel.XLWorkbook

# Sheet 1: Sales Data
$salesSheet = $workbook.Worksheets.Add("Sales")
$salesSheet.Cell(1, 1).Value = "OrderID"
$salesSheet.Cell(1, 2).Value = "Product"
$salesSheet.Cell(1, 3).Value = "Region"
$salesSheet.Cell(1, 4).Value = "Amount"
$salesSheet.Cell(1, 5).Value = "Date"

# Add sample sales data
$salesData = @(
    @(1001, "Laptop", "North", 1200, "2024-01-15"),
    @(1002, "Mouse", "South", 25, "2024-01-16"),
    @(1003, "Keyboard", "East", 75, "2024-01-17"),
    @(1004, "Monitor", "West", 300, "2024-01-18"),
    @(1005, "Laptop", "North", 1200, "2024-01-19"),
    @(1006, "Headphones", "South", 80, "2024-01-20"),
    @(1007, "Mouse", "East", 25, "2024-01-21"),
    @(1008, "Laptop", "West", 1200, "2024-01-22"),
    @(1009, "Monitor", "North", 300, "2024-01-23"),
    @(1010, "Keyboard", "South", 75, "2024-01-24")
)

$row = 2
foreach ($sale in $salesData) {
    $salesSheet.Cell($row, 1).Value = $sale[0]
    $salesSheet.Cell($row, 2).Value = $sale[1]
    $salesSheet.Cell($row, 3).Value = $sale[2]
    $salesSheet.Cell($row, 4).Value = $sale[3]
    $salesSheet.Cell($row, 5).Value = $sale[4]
    $row++
}

# Format as table
$salesRange = $salesSheet.Range(1, 1, $row - 1, 5)
$salesTable = $salesRange.CreateTable("SalesTable")

# Sheet 2: Inventory
$inventorySheet = $workbook.Worksheets.Add("Inventory")
$inventorySheet.Cell(1, 1).Value = "ProductID"
$inventorySheet.Cell(1, 2).Value = "Product"
$inventorySheet.Cell(1, 3).Value = "Stock"
$inventorySheet.Cell(1, 4).Value = "Warehouse"

$inventoryData = @(
    @(101, "Laptop", 45, "Main"),
    @(102, "Mouse", 250, "Main"),
    @(103, "Keyboard", 120, "Main"),
    @(104, "Monitor", 60, "Backup"),
    @(105, "Headphones", 180, "Main")
)

$row = 2
foreach ($item in $inventoryData) {
    $inventorySheet.Cell($row, 1).Value = $item[0]
    $inventorySheet.Cell($row, 2).Value = $item[1]
    $inventorySheet.Cell($row, 3).Value = $item[2]
    $inventorySheet.Cell($row, 4).Value = $item[3]
    $row++
}

# Format as table
$inventoryRange = $inventorySheet.Range(1, 1, $row - 1, 4)
$inventoryTable = $inventoryRange.CreateTable("InventoryTable")

# Sheet 3: Returns
$returnsSheet = $workbook.Worksheets.Add("Returns")
$returnsSheet.Cell(1, 1).Value = "ReturnID"
$returnsSheet.Cell(1, 2).Value = "Product"
$returnsSheet.Cell(1, 3).Value = "Reason"
$returnsSheet.Cell(1, 4).Value = "Date"

$returnsData = @(
    @(5001, "Laptop", "Defective", "2024-01-20"),
    @(5002, "Mouse", "Changed Mind", "2024-01-22"),
    @(5003, "Monitor", "Defective", "2024-01-25")
)

$row = 2
foreach ($return in $returnsData) {
    $returnsSheet.Cell($row, 1).Value = $return[0]
    $returnsSheet.Cell($row, 2).Value = $return[1]
    $returnsSheet.Cell($row, 3).Value = $return[2]
    $returnsSheet.Cell($row, 4).Value = $return[3]
    $row++
}

# Format as table
$returnsRange = $returnsSheet.Range(1, 1, $row - 1, 4)
$returnsTable = $returnsRange.CreateTable("ReturnsTable")

# Save workbook
$fullPath = Resolve-Path $OutputPath -ErrorAction SilentlyContinue
if (-not $fullPath) {
    $fullPath = Join-Path (Get-Location) $OutputPath
}

$workbook.SaveAs($fullPath)
$workbook.Dispose()

Write-Host "✓ Sample workbook created: $fullPath" -ForegroundColor Green
Write-Host "  - Sales sheet with 10 orders (SalesTable)"
Write-Host "  - Inventory sheet with 5 products (InventoryTable)"
Write-Host "  - Returns sheet with 3 returns (ReturnsTable)"
