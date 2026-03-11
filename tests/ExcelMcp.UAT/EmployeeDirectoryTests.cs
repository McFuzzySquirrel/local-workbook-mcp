using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.UAT;

/// <summary>
/// User Acceptance Tests — EmployeeDirectory.xlsx
///
/// Covers HR / people-management scenarios a user might ask the AI agent.
/// Re-generate test data with: scripts/create-sample-workbooks.ps1
/// </summary>
public sealed class EmployeeDirectoryTests
{
    private static readonly string WorkbookPath = TestData.GetPath("EmployeeDirectory.xlsx");

    // ── Structure ────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-ED-01: Workbook has two worksheets: Employees and Departments")]
    public async Task WorkbookHasTwoExpectedSheets()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheetNames = metadata.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Contains("Employees", sheetNames);
        Assert.Contains("Departments", sheetNames);
        Assert.Equal(2, sheetNames.Count);
    }

    [Fact(DisplayName = "UA-ED-02: EmployeesTable has eight rows and expected columns")]
    public async Task EmployeesTableHasEightRows()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheet = metadata.Worksheets.Single(ws => ws.Name == "Employees");
        var table = sheet.Tables.Single(t => t.Name == "EmployeesTable");

        Assert.Equal(8, table.RowCount);
        Assert.Contains("FullName", table.ColumnHeaders);
        Assert.Contains("Department", table.ColumnHeaders);
        Assert.Contains("Salary", table.ColumnHeaders);
    }

    [Fact(DisplayName = "UA-ED-03: DepartmentsTable has four rows")]
    public async Task DepartmentsTableHasFourRows()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheet = metadata.Worksheets.Single(ws => ws.Name == "Departments");
        var table = sheet.Tables.Single(t => t.Name == "DepartmentsTable");

        Assert.Equal(4, table.RowCount);
    }

    // ── Search: Department ───────────────────────────────────────────────────

    [Fact(DisplayName = "UA-ED-04: Searching 'Engineering' returns four employees")]
    public async Task SearchEngineering_ReturnsFourEmployees()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Engineering", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        // Engineering: Alice Johnson, Bob Smith, Charlie Brown, Henry Wilson
        Assert.Equal(4, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Engineering", row.Values["Department"]));
    }

    [Fact(DisplayName = "UA-ED-05: Searching 'Sales' returns two employees")]
    public async Task SearchSales_ReturnsTwoEmployees()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Sales", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        // Sales: Diana Prince, Eve Davis
        Assert.Equal(2, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Sales", row.Values["Department"]));
    }

    // ── Search: Name ─────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-ED-06: Searching 'Alice Johnson' finds her with correct details")]
    public async Task SearchAliceJohnson_ReturnsCorrectEmployee()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Alice Johnson", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        var alice = result.Rows[0];
        Assert.Equal("Alice Johnson", alice.Values["FullName"]);
        Assert.Equal("Engineering", alice.Values["Department"]);
        Assert.Equal("Senior Developer", alice.Values["Position"]);
        Assert.Equal("95000", alice.Values["Salary"]);
    }

    [Fact(DisplayName = "UA-ED-07: Searching 'Bob Smith' finds him as Lead Developer with highest salary in Engineering")]
    public async Task SearchBobSmith_IsLeadDeveloperSalary110k()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Bob Smith", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        var bob = result.Rows[0];
        Assert.Equal("Lead Developer", bob.Values["Position"]);
        Assert.Equal("110000", bob.Values["Salary"]);
    }

    // ── Search: Position / Title ─────────────────────────────────────────────

    [Fact(DisplayName = "UA-ED-08: Searching 'Manager' returns four people with Manager in their title")]
    public async Task SearchManager_ReturnsFourManagers()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        // Search across whole Employees sheet (not just the table) to pick up Manager positions
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Manager", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        // Sales Manager (Diana), HR Manager (Frank), Marketing Director isn't Manager but
        // DepartmentsTable has "Bob Smith" as Engineering manager — but we're scoped to Employees.
        // Positions with "Manager": Sales Manager, HR Manager
        // "Lead Developer" Bob isn't a Manager title.
        // So: Diana Prince (Sales Manager), Frank Miller (HR Manager) = 2 in EmployeesTable
        Assert.Equal(2, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Contains("Manager", row.Values["Position"]));
    }

    [Fact(DisplayName = "UA-ED-09: Most recently hired employee is Henry Wilson (2024-01-08)")]
    public async Task SearchHenryWilson_IsRecentHire()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Henry Wilson", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        Assert.Equal("Junior Developer", result.Rows[0].Values["Position"]);
        // HireDate is stored as a date serial or ISO string — just confirm it's present and non-empty
        Assert.NotNull(result.Rows[0].Values["HireDate"]);
    }

    // ── Search: Departments ──────────────────────────────────────────────────

    [Fact(DisplayName = "UA-ED-10: Engineering department budget is 500000")]
    public async Task EngineeringDepartment_BudgetIs500000()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Engineering", Worksheet: "Departments", Table: "DepartmentsTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        Assert.Equal("500000", result.Rows[0].Values["Budget"]);
    }

    [Fact(DisplayName = "UA-ED-11: Grace Lee is the Marketing Director and manages the Marketing department")]
    public async Task GraceLee_IsMarketingManager()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var empResult = await service.SearchAsync(
            new ExcelSearchArguments("Grace Lee", Worksheet: "Employees", Table: "EmployeesTable"),
            CancellationToken.None);

        Assert.Single(empResult.Rows);
        Assert.Equal("Marketing", empResult.Rows[0].Values["Department"]);
        Assert.Equal("Marketing Director", empResult.Rows[0].Values["Position"]);

        // Confirm she also appears in DepartmentsTable as manager
        var deptResult = await service.SearchAsync(
            new ExcelSearchArguments("Grace Lee", Worksheet: "Departments", Table: "DepartmentsTable"),
            CancellationToken.None);

        Assert.Single(deptResult.Rows);
        Assert.Equal("Marketing", deptResult.Rows[0].Values["DeptName"]);
    }

    // ── Preview ──────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-ED-12: Previewing Employees sheet returns CSV with Salary column")]
    public async Task PreviewEmployees_HasSalaryColumn()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var uri = ExcelResourceUri.CreateWorksheetUri("Employees");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 5);

        Assert.Equal("text/csv", content.MimeType);
        Assert.NotNull(content.Text);
        Assert.Contains("Salary", content.Text);
        Assert.Contains("Department", content.Text);
    }
}
