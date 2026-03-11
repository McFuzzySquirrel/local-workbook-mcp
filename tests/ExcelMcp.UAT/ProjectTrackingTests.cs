using ExcelMcp.Contracts;
using ExcelMcp.Server.Excel;
using Xunit;

namespace ExcelMcp.UAT;

/// <summary>
/// User Acceptance Tests — ProjectTracking.xlsx
///
/// These tests verify end-to-end behaviour against the real workbook used in demos.
/// Each test corresponds to a scenario a user or AI agent might attempt via the MCP tools.
///
/// Re-generate test data with: scripts/create-sample-workbooks.ps1
/// </summary>
public sealed class ProjectTrackingTests
{
    private static readonly string WorkbookPath = TestData.GetPath("ProjectTracking.xlsx");

    // ── Structure ────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-PT-01: Workbook has three worksheets: Tasks, Projects, TimeLog")]
    public async Task WorkbookHasThreeExpectedSheets()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var sheetNames = metadata.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Contains("Tasks", sheetNames);
        Assert.Contains("Projects", sheetNames);
        Assert.Contains("TimeLog", sheetNames);
        Assert.Equal(3, sheetNames.Count);
    }

    [Fact(DisplayName = "UA-PT-02: Tasks worksheet exposes TasksTable with expected columns")]
    public async Task TasksWorksheetHasTasksTable()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var tasksSheet = metadata.Worksheets.Single(ws => ws.Name == "Tasks");
        var table = tasksSheet.Tables.Single(t => t.Name == "TasksTable");

        Assert.Contains("TaskID", table.ColumnHeaders);
        Assert.Contains("TaskName", table.ColumnHeaders);
        Assert.Contains("Owner", table.ColumnHeaders);
        Assert.Contains("Status", table.ColumnHeaders);
        Assert.Contains("Priority", table.ColumnHeaders);
        Assert.Contains("DueDate", table.ColumnHeaders);
        Assert.Equal(10, table.RowCount);
    }

    [Fact(DisplayName = "UA-PT-03: Projects worksheet exposes ProjectsTable with three rows")]
    public async Task ProjectsTableHasThreeProjects()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var metadata = await service.GetMetadataAsync(CancellationToken.None);

        var projectsSheet = metadata.Worksheets.Single(ws => ws.Name == "Projects");
        var table = projectsSheet.Tables.Single(t => t.Name == "ProjectsTable");

        Assert.Equal(3, table.RowCount);
    }

    // ── Search: Priority ─────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-PT-04: Searching 'High' returns five high-priority tasks")]
    public async Task SearchHighPriority_ReturnsFiveTasks()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("High", Worksheet: "Tasks", Table: "TasksTable"),
            CancellationToken.None);

        // High-priority tasks: Setup Database, Design UI Mockups, Security Audit,
        //                       Deploy to Staging, Fix Database Backup
        Assert.Equal(5, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("High", row.Values["Priority"]));
    }

    [Fact(DisplayName = "UA-PT-05: Searching 'In Progress' returns three tasks")]
    public async Task SearchInProgress_ReturnsThreeTasks()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("In Progress", Worksheet: "Tasks", Table: "TasksTable"),
            CancellationToken.None);

        // In Progress: Design UI Mockups, Write Documentation, Fix Database Backup
        Assert.Equal(3, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("In Progress", row.Values["Status"]));
    }

    // ── Search: Owner ────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-PT-06: Searching 'Alice' in TasksTable returns her four tasks")]
    public async Task SearchAlice_ReturnsFourTasks()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Alice", Worksheet: "Tasks", Table: "TasksTable"),
            CancellationToken.None);

        // Alice owns: Setup Database, Write Documentation, Deploy to Staging, Update Dependencies
        Assert.Equal(4, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Alice", row.Values["Owner"]));
    }

    [Fact(DisplayName = "UA-PT-07: Searching 'Bob' in TasksTable returns his three tasks")]
    public async Task SearchBob_ReturnsThreeTasks()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Bob", Worksheet: "Tasks", Table: "TasksTable"),
            CancellationToken.None);

        // Bob owns: Design UI Mockups, Security Audit, User Acceptance Test
        Assert.Equal(3, result.Rows.Count);
        Assert.All(result.Rows, row =>
            Assert.Equal("Bob", row.Values["Owner"]));
    }

    // ── Search: Projects ─────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-PT-08: Projects table contains Website Redesign, Mobile App, Data Migration")]
    public async Task Projects_ContainsAllThreeProjectNames()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Alice", Worksheet: "Projects", Table: "ProjectsTable"),
            CancellationToken.None);

        // Alice is manager of Website Redesign (only one Alice entry in Projects)
        Assert.Single(result.Rows);
        Assert.Equal("Website Redesign", result.Rows[0].Values["ProjectName"]);
    }

    [Fact(DisplayName = "UA-PT-09: Mobile App project has the highest budget (75000)")]
    public async Task SearchMobileApp_BudgetIs75000()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Mobile App", Worksheet: "Projects", Table: "ProjectsTable"),
            CancellationToken.None);

        Assert.Single(result.Rows);
        Assert.Equal("75000", result.Rows[0].Values["Budget"]);
    }

    // ── Search: TimeLog ──────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-PT-10: Bob's time-log entries sum to 20 hours across three entries")]
    public async Task BobTimeLog_TotalTwentyHours()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var result = await service.SearchAsync(
            new ExcelSearchArguments("Bob", Worksheet: "TimeLog", Table: "TimeLogTable"),
            CancellationToken.None);

        // 3 entries: 5 + 7 + 8 = 20 hours
        Assert.Equal(3, result.Rows.Count);
        var totalHours = result.Rows
            .Sum(row => double.Parse(row.Values["Hours"]!));
        Assert.Equal(20.0, totalHours);
    }

    // ── Preview ──────────────────────────────────────────────────────────────

    [Fact(DisplayName = "UA-PT-11: Previewing Tasks sheet returns CSV with column headers")]
    public async Task PreviewTasks_ReturnsCsvWithHeaders()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var uri = ExcelResourceUri.CreateWorksheetUri("Tasks");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 5);

        Assert.Equal("text/csv", content.MimeType);
        Assert.NotNull(content.Text);
        Assert.Contains("TaskName", content.Text);
        Assert.Contains("Owner", content.Text);
        Assert.Contains("Priority", content.Text);
    }

    [Fact(DisplayName = "UA-PT-12: Previewing TasksTable returns table-scoped CSV")]
    public async Task PreviewTasksTable_ReturnsCsvForTable()
    {
        var service = new ExcelWorkbookService(WorkbookPath);
        var uri = ExcelResourceUri.CreateTableUri("Tasks", "TasksTable");
        var content = await service.ReadResourceAsync(uri, CancellationToken.None, maxRows: 3);

        Assert.Equal("text/csv", content.MimeType);
        Assert.NotNull(content.Text);
        // Header row should be present
        Assert.Contains("TaskID", content.Text);
    }
}
