#!/usr/bin/env pwsh
# Create a test workbook with pivot table for testing pivot analysis feature

Add-Type -Path "D:\GitHub Projects\local-workbook-mcp\src\ExcelMcp.Server\bin\Debug\net9.0\ClosedXML.dll"

$workbookPath = Join-Path $PSScriptRoot "..\test-data\SalesWithPivot.xlsx"
$workbook = New-Object ClosedXML.Excel.XLWorkbook

# Create source data sheet
$dataSheet = $workbook.Worksheets.Add("SalesData")
$dataSheet.Cell("A1").Value = "Region"
$dataSheet.Cell("B1").Value = "Product"
$dataSheet.Cell("C1").Value = "SalesPerson"
$dataSheet.Cell("D1").Value = "Amount"
$dataSheet.Cell("E1").Value = "Quarter"

# Add sample data
$data = @(
    @("East", "Widget", "John", 1200, "Q1"),
    @("East", "Widget", "Sarah", 1500, "Q1"),
    @("West", "Gadget", "Mike", 2000, "Q1"),
    @("West", "Widget", "John", 1800, "Q2"),
    @("East", "Gadget", "Sarah", 2200, "Q2"),
    @("West", "Widget", "Mike", 1600, "Q2"),
    @("East", "Widget", "John", 1900, "Q3"),
    @("West", "Gadget", "Sarah", 2100, "Q3"),
    @("East", "Gadget", "Mike", 1700, "Q3"),
    @("West", "Widget", "John", 2300, "Q4")
)

$row = 2
foreach ($record in $data) {
    $dataSheet.Cell($row, 1).Value = $record[0]
    $dataSheet.Cell($row, 2).Value = $record[1]
    $dataSheet.Cell($row, 3).Value = $record[2]
    $dataSheet.Cell($row, 4).Value = $record[3]
    $dataSheet.Cell($row, 5).Value = $record[4]
    $row++
}

# Create table from the data
$tableRange = $dataSheet.Range("A1:E$($data.Count + 1)")
$table = $tableRange.CreateTable("SalesTable")

# Create pivot table sheet
$pivotSheet = $workbook.Worksheets.Add("SalesPivot")
$pivotTable = $pivotSheet.PivotTables.Add("SalesSummary", $pivotSheet.Cell("A1"), $table.AsRange())

# Configure pivot table
$pivotTable.RowLabels.Add("Region") | Out-Null
$pivotTable.RowLabels.Add("Product") | Out-Null
$pivotTable.ColumnLabels.Add("Quarter") | Out-Null
$pivotTable.Values.Add("Amount").SetSummaryFormula([ClosedXML.Excel.XLPivotSummary]::Sum) | Out-Null

$workbook.SaveAs($workbookPath)
$workbook.Dispose()

Write-Host "✓ Created workbook with pivot table at: $workbookPath" -ForegroundColor Green
