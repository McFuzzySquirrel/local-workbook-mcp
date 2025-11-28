using ExcelMcp.Contracts;
using Xunit;

namespace ExcelMcp.Contracts.Tests;

public sealed class WriteBackContractsTests
{
    [Fact]
    public void UpdateCellArguments_DefaultValues()
    {
        var args = new UpdateCellArguments("Sheet1", "A1", "NewValue");

        Assert.Equal("Sheet1", args.Worksheet);
        Assert.Equal("A1", args.CellAddress);
        Assert.Equal("NewValue", args.Value);
        Assert.Null(args.Reason);
    }

    [Fact]
    public void UpdateCellArguments_WithReason()
    {
        var args = new UpdateCellArguments(
            Worksheet: "Data",
            CellAddress: "B5",
            Value: "100",
            Reason: "Corrected calculation error");

        Assert.Equal("Data", args.Worksheet);
        Assert.Equal("B5", args.CellAddress);
        Assert.Equal("100", args.Value);
        Assert.Equal("Corrected calculation error", args.Reason);
    }

    [Fact]
    public void UpdateCellResult_CanBeCreated()
    {
        var result = new UpdateCellResult(
            Worksheet: "Sheet1",
            CellAddress: "A1",
            PreviousValue: "Old",
            NewValue: "New",
            Timestamp: DateTimeOffset.UtcNow,
            AuditId: "audit-001");

        Assert.Equal("Sheet1", result.Worksheet);
        Assert.Equal("A1", result.CellAddress);
        Assert.Equal("Old", result.PreviousValue);
        Assert.Equal("New", result.NewValue);
        Assert.NotNull(result.AuditId);
    }

    [Fact]
    public void UpdateCellResult_NullPreviousValue()
    {
        var result = new UpdateCellResult(
            Worksheet: "Sheet1",
            CellAddress: "A1",
            PreviousValue: null,
            NewValue: "New",
            Timestamp: DateTimeOffset.UtcNow,
            AuditId: null);

        Assert.Null(result.PreviousValue);
        Assert.Null(result.AuditId);
    }

    [Fact]
    public void AddWorksheetArguments_DefaultValues()
    {
        var args = new AddWorksheetArguments("NewSheet");

        Assert.Equal("NewSheet", args.Name);
        Assert.Null(args.Position);
        Assert.Null(args.Reason);
    }

    [Fact]
    public void AddWorksheetArguments_AllParameters()
    {
        var args = new AddWorksheetArguments(
            Name: "Analysis",
            Position: 3,
            Reason: "Adding analysis worksheet");

        Assert.Equal("Analysis", args.Name);
        Assert.Equal(3, args.Position);
        Assert.Equal("Adding analysis worksheet", args.Reason);
    }

    [Fact]
    public void AddWorksheetResult_CanBeCreated()
    {
        var result = new AddWorksheetResult(
            Name: "NewSheet",
            Position: 1,
            Timestamp: DateTimeOffset.UtcNow,
            AuditId: "audit-002");

        Assert.Equal("NewSheet", result.Name);
        Assert.Equal(1, result.Position);
        Assert.NotNull(result.AuditId);
    }

    [Fact]
    public void AddAnnotationArguments_DefaultValues()
    {
        var args = new AddAnnotationArguments("Sheet1", "A1", "This is a note");

        Assert.Equal("Sheet1", args.Worksheet);
        Assert.Equal("A1", args.CellAddress);
        Assert.Equal("This is a note", args.Text);
        Assert.Null(args.Author);
    }

    [Fact]
    public void AddAnnotationArguments_WithAuthor()
    {
        var args = new AddAnnotationArguments(
            Worksheet: "Data",
            CellAddress: "C10",
            Text: "Review this value",
            Author: "Agent");

        Assert.Equal("Data", args.Worksheet);
        Assert.Equal("C10", args.CellAddress);
        Assert.Equal("Review this value", args.Text);
        Assert.Equal("Agent", args.Author);
    }

    [Fact]
    public void AddAnnotationResult_CanBeCreated()
    {
        var result = new AddAnnotationResult(
            Worksheet: "Sheet1",
            CellAddress: "A1",
            Text: "Annotation text",
            Author: "Agent",
            Timestamp: DateTimeOffset.UtcNow,
            AuditId: "audit-003");

        Assert.Equal("Sheet1", result.Worksheet);
        Assert.Equal("A1", result.CellAddress);
        Assert.Equal("Annotation text", result.Text);
        Assert.Equal("Agent", result.Author);
        Assert.NotNull(result.AuditId);
    }

    [Fact]
    public void AuditEntry_CanBeCreated()
    {
        var details = new Dictionary<string, string?>
        {
            { "worksheet", "Sheet1" },
            { "cell", "A1" },
            { "oldValue", "Old" },
            { "newValue", "New" }
        };

        var entry = new AuditEntry(
            Id: "audit-001",
            OperationType: "UpdateCell",
            Description: "Updated cell A1 in Sheet1",
            Timestamp: DateTimeOffset.UtcNow,
            Reason: "Correction",
            Details: details);

        Assert.Equal("audit-001", entry.Id);
        Assert.Equal("UpdateCell", entry.OperationType);
        Assert.Equal("Updated cell A1 in Sheet1", entry.Description);
        Assert.Equal("Correction", entry.Reason);
        Assert.Equal(4, entry.Details.Count);
    }

    [Fact]
    public void AuditEntry_EmptyDetails()
    {
        var entry = new AuditEntry(
            Id: "audit-002",
            OperationType: "ListStructure",
            Description: "Listed workbook structure",
            Timestamp: DateTimeOffset.UtcNow,
            Reason: null,
            Details: new Dictionary<string, string?>());

        Assert.Empty(entry.Details);
        Assert.Null(entry.Reason);
    }

    [Fact]
    public void GetAuditTrailArguments_DefaultValues()
    {
        var args = new GetAuditTrailArguments();

        Assert.Null(args.Since);
        Assert.Null(args.Until);
        Assert.Null(args.OperationType);
        Assert.Null(args.Limit);
    }

    [Fact]
    public void GetAuditTrailArguments_AllParameters()
    {
        var since = DateTimeOffset.UtcNow.AddDays(-7);
        var until = DateTimeOffset.UtcNow;

        var args = new GetAuditTrailArguments(
            Since: since,
            Until: until,
            OperationType: "UpdateCell",
            Limit: 50);

        Assert.Equal(since, args.Since);
        Assert.Equal(until, args.Until);
        Assert.Equal("UpdateCell", args.OperationType);
        Assert.Equal(50, args.Limit);
    }

    [Fact]
    public void GetAuditTrailResult_CanBeCreated()
    {
        var entries = new[]
        {
            new AuditEntry(
                "audit-001",
                "UpdateCell",
                "Updated A1",
                DateTimeOffset.UtcNow,
                null,
                new Dictionary<string, string?>())
        };

        var result = new GetAuditTrailResult(entries, 1);

        Assert.Single(result.Entries);
        Assert.Equal(1, result.TotalCount);
    }

    [Fact]
    public void GetAuditTrailResult_EmptyResult()
    {
        var result = new GetAuditTrailResult(Array.Empty<AuditEntry>(), 0);

        Assert.Empty(result.Entries);
        Assert.Equal(0, result.TotalCount);
    }

    [Fact]
    public void UpdateCellArguments_RecordEquality()
    {
        var args1 = new UpdateCellArguments("Sheet1", "A1", "Value", "Reason");
        var args2 = new UpdateCellArguments("Sheet1", "A1", "Value", "Reason");

        Assert.Equal(args1, args2);
    }

    [Fact]
    public void AddWorksheetArguments_RecordEquality()
    {
        var args1 = new AddWorksheetArguments("Sheet", 1, "Reason");
        var args2 = new AddWorksheetArguments("Sheet", 1, "Reason");

        Assert.Equal(args1, args2);
    }

    [Fact]
    public void AddAnnotationArguments_RecordEquality()
    {
        var args1 = new AddAnnotationArguments("Sheet1", "A1", "Text", "Author");
        var args2 = new AddAnnotationArguments("Sheet1", "A1", "Text", "Author");

        Assert.Equal(args1, args2);
    }
}
