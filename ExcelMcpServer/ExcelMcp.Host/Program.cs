using ExcelMcp.Core;
using System.Text.Json;
using System.Net.WebSockets;
using System.Text;

var builder = WebApplication.CreateBuilder(args);
builder.Configuration.AddJsonFile("appsettings.json", optional: true);

var app = builder.Build();

// configuration: path to excel file
var excelFile = builder.Configuration.GetValue<string>("ExcelFile") ?? "data.xlsx";
if (!File.Exists(excelFile))
{
    Console.WriteLine($"Warning: Excel file not found: {excelFile}");
}

// Initialize workbook cache for on-demand diffs
WorkbookCache.Initialize(excelFile);

// in-memory clients
var sockets = new List<WebSocket>();

// Create and start monitor
var monitor = new ExcelDiffMonitor(excelFile);
monitor.OnChange += async ev =>
{
    var json = JsonSerializer.Serialize(ev);
    var buffer = Encoding.UTF8.GetBytes(json);
    var segment = new ArraySegment<byte>(buffer);

    // broadcast to all open sockets
    var closed = new List<WebSocket>();
    foreach (var ws in sockets)
    {
        if (ws.State == WebSocketState.Open)
        {
            try
            {
                await ws.SendAsync(segment, WebSocketMessageType.Text, true, CancellationToken.None);
            }
            catch
            {
                closed.Add(ws);
            }
        }
        else closed.Add(ws);
    }
    // remove closed sockets
    closed.ForEach(c => sockets.Remove(c));
};

monitor.Start();

// REST: tools
app.MapGet("/mcp/tools", () =>
{
    object[] tools = new object[]
    {
        new { name = "listSheets", description = "Return sheet names", inputs = (object?)null, outputs = new { sheets = "string[]" } },
        new { name = "listTables", description = "Return table names for sheet", inputs = new { sheetName = "string" }, outputs = new { tables = "string[]" } },
    new { name = "getSheetData", description = "Sheet data matrix (raw cells)", inputs = new { sheetName = "string" }, outputs = new { data = "object[][]" } },
    new { name = "getTableData", description = "Table data", inputs = new { sheetName = "string", tableName = "string" }, outputs = new { data = "object[][]" } },
    new { name = "setCell", description = "Set a single cell by A1 address", inputs = new { sheetName = "string", address = "string", value = "string" }, outputs = new { ok = "bool" } },
    new { name = "appendRow", description = "Append a row to the end of the sheet", inputs = new { sheetName = "string", values = "string[]" }, outputs = new { ok = "bool" } }
    };
    return Results.Json(tools);
});
 
// Tool invocation endpoint
app.MapPost("/mcp/tools/invoke", (ToolInvokeRequest req) =>
{
    if (req is null || string.IsNullOrWhiteSpace(req.name))
        return Results.BadRequest("Missing tool name");

    var args = req.args ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    var data = ExcelReader.LoadWorkbook(excelFile);

    switch (req.name)
    {
        case "listSheets":
            return Results.Json(new { sheets = data.Keys.ToArray() });

        case "listTables":
            if (!args.TryGetValue("sheetName", out var sheetName) || string.IsNullOrWhiteSpace(sheetName))
                return Results.BadRequest("sheetName is required");
            if (!data.TryGetValue(sheetName, out var tables))
                return Results.NotFound($"Sheet {sheetName} not found");
            return Results.Json(new { tables = tables.Select(t => t.TableName).ToArray() });

        case "getSheetData":
            if (!args.TryGetValue("sheetName", out var sname) || string.IsNullOrWhiteSpace(sname))
                return Results.BadRequest("sheetName is required");
            try
            {
                var matrix = ExcelMcp.Core.ExcelReadHelpers.ReadSheetAsMatrix(excelFile, sname);
                return Results.Json(new { data = matrix });
            }
            catch (Exception ex)
            {
                return Results.Problem(ex.Message, statusCode: 500);
            }

        case "getTableData":
            if (!args.TryGetValue("sheetName", out var s2) || string.IsNullOrWhiteSpace(s2))
                return Results.BadRequest("sheetName is required");
            if (!args.TryGetValue("tableName", out var tname) || string.IsNullOrWhiteSpace(tname))
                return Results.BadRequest("tableName is required");
            if (!data.TryGetValue(s2, out var tbls))
                return Results.NotFound($"Sheet {s2} not found");
            var tbl = tbls.FirstOrDefault(t => string.Equals(t.TableName, tname, StringComparison.OrdinalIgnoreCase));
            if (tbl is null) return Results.NotFound($"Table {tname} not found in sheet {s2}");
            var rows = tbl.Rows.Cast<System.Data.DataRow>().Select(r => tbl.Columns.Cast<System.Data.DataColumn>().Select(c => new { column = c.ColumnName, value = r[c]?.ToString() }).ToArray());
            return Results.Json(new { data = rows });

        case "setCell":
            {
                if (!args.TryGetValue("sheetName", out var sname2) || string.IsNullOrWhiteSpace(sname2))
                    return Results.BadRequest("sheetName is required");
                if (!args.TryGetValue("address", out var address) || string.IsNullOrWhiteSpace(address))
                    return Results.BadRequest("address is required");
                args.TryGetValue("value", out var val);
                try
                {
                    ExcelWriteHelpers.WithRetry(() => ExcelWriteHelpers.SetCellValue(excelFile, sname2, address, val ?? string.Empty));
                    return Results.Json(new { ok = true });
                }
                catch (Exception ex)
                {
                    return Results.Problem(ex.Message, statusCode: 500);
                }
            }

        case "appendRow":
            {
                if (!args.TryGetValue("sheetName", out var sname3) || string.IsNullOrWhiteSpace(sname3))
                    return Results.BadRequest("sheetName is required");
                if (!req.args!.TryGetValue("values", out var _)) { /* placeholder to satisfy analyzer */ }
                // For POST form, values will come via a JSON array in a different shape; support alternate body binder later.
                // For now, accept comma-separated 'values' in args (simple agent usage):
                var valsCsv = args.TryGetValue("values", out var v) ? v : null;
                if (string.IsNullOrWhiteSpace(valsCsv)) return Results.BadRequest("values (comma-separated) is required");
                var values = valsCsv.Split(',').Select(x => x.Trim()).ToArray();
                try
                {
                    ExcelWriteHelpers.WithRetry(() => ExcelWriteHelpers.AppendRow(excelFile, sname3, values));
                    return Results.Json(new { ok = true });
                }
                catch (Exception ex)
                {
                    return Results.Problem(ex.Message, statusCode: 500);
                }
            }

        default:
            return Results.NotFound($"Unknown tool: {req.name}");
    }
});
 
// Tool invocation endpoints (simple MCP-style schema)
app.MapPost("/mcp/tools/listSheets", () =>
{
    var data = ExcelReader.LoadWorkbook(excelFile);
    var sheets = data.Keys.ToArray();
    return Results.Json(new { sheets });
});

app.MapPost("/mcp/tools/listTables", async (HttpRequest req) =>
{
    var payload = await System.Text.Json.JsonSerializer.DeserializeAsync<ListTablesRequest>(req.Body);
    if (payload is null || string.IsNullOrWhiteSpace(payload.sheetName)) return Results.BadRequest("sheetName is required");
    var data = ExcelReader.LoadWorkbook(excelFile);
    if (!data.ContainsKey(payload.sheetName)) return Results.NotFound($"Sheet {payload.sheetName} not found");
    var tables = data[payload.sheetName].Select(t => t.TableName).ToArray();
    return Results.Json(new { tables });
});

app.MapPost("/mcp/tools/getSheetData", async (HttpRequest req) =>
{
    var payload = await System.Text.Json.JsonSerializer.DeserializeAsync<GetSheetDataRequest>(req.Body);
    if (payload is null || string.IsNullOrWhiteSpace(payload.sheetName)) return Results.BadRequest("sheetName is required");
    try
    {
        var matrix = ExcelMcp.Core.ExcelReadHelpers.ReadSheetAsMatrix(excelFile, payload.sheetName);
        return Results.Json(new { data = matrix });
    }
    catch (KeyNotFoundException)
    {
        return Results.NotFound($"Sheet {payload.sheetName} not found");
    }
    catch (Exception ex)
    {
        return Results.Problem(ex.Message, statusCode: 500);
    }
});

app.MapPost("/mcp/tools/getTableData", async (HttpRequest req) =>
{
    var payload = await System.Text.Json.JsonSerializer.DeserializeAsync<GetTableDataRequest>(req.Body);
    if (payload is null || string.IsNullOrWhiteSpace(payload.sheetName) || string.IsNullOrWhiteSpace(payload.tableName))
        return Results.BadRequest("sheetName and tableName are required");
    var data = ExcelReader.LoadWorkbook(excelFile);
    if (!data.ContainsKey(payload.sheetName)) return Results.NotFound($"Sheet {payload.sheetName} not found");
    var tbl = data[payload.sheetName].FirstOrDefault(t => string.Equals(t.TableName, payload.tableName, StringComparison.OrdinalIgnoreCase));
    if (tbl == null) return Results.NotFound($"Table {payload.tableName} not found in sheet {payload.sheetName}");
    var rows = tbl.Rows.Cast<System.Data.DataRow>().Select(r => r.ItemArray.Select(c => c?.ToString()).ToArray()).ToArray();
    return Results.Json(new { data = rows });
});

// REST: resources (sheets and tables)
app.MapGet("/mcp/resources", () =>
{
    var data = ExcelReader.LoadWorkbook(excelFile);
    var resources = new List<object>();
    foreach (var sheet in data)
    {
        var sheetUri = $"excel://{Path.GetFileName(excelFile)}/{Uri.EscapeDataString(sheet.Key)}";
        resources.Add(new { uri = sheetUri, type = "sheet" });
        for (int i = 0; i < sheet.Value.Count; i++)
        {
            var table = sheet.Value[i];
            var tableUri = $"{sheetUri}/{Uri.EscapeDataString(table.TableName)}";
            resources.Add(new { uri = tableUri, type = "table", columns = table.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToArray() });
        }
    }
    return Results.Json(resources);
});

// On-demand diff endpoint
app.MapGet("/mcp/diff", () =>
{
    var (diffs, ts) = WorkbookCache.ReloadAndDiff();
    return Results.Json(new { timestampUtc = ts, count = diffs.Count, diffs });
});

// REST: query sheet/table
app.MapGet("/mcp/query", (string sheet, string? table) =>
{
    var data = ExcelReader.LoadWorkbook(excelFile);
    if (!data.ContainsKey(sheet)) return Results.NotFound($"Sheet {sheet} not found");
    if (string.IsNullOrEmpty(table))
    {
        // return raw matrix (rows of arrays)
        var d = data[sheet].SelectMany(t => t.Rows.Cast<System.Data.DataRow>().Select(r => r.ItemArray.Select(c => c?.ToString()).ToArray())).ToArray();
        return Results.Json(d);
    }
    else
    {
        var tbl = data[sheet].FirstOrDefault(t => string.Equals(t.TableName, table, StringComparison.OrdinalIgnoreCase));
        if (tbl == null) return Results.NotFound($"Table {table} not found in sheet {sheet}");
        var rows = tbl.Rows.Cast<System.Data.DataRow>().Select(r => tbl.Columns.Cast<System.Data.DataColumn>().Select(c => new { column = c.ColumnName, value = r[c]?.ToString() }).ToArray());
        return Results.Json(rows);
    }
});

// REST: workbook dump
app.MapGet("/workbook", () =>
{
    var data = ExcelReader.LoadWorkbook(excelFile);
    // serialize to simplified JSON structure
    var outObj = data.ToDictionary(
        kv => kv.Key,
        kv => kv.Value.Select(t => new { tableName = t.TableName, columns = t.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName).ToArray(), rows = t.Rows.Cast<System.Data.DataRow>().Select(r => r.ItemArray.Select(c => c?.ToString()).ToArray()).ToArray() }).ToArray()
    );
    return Results.Json(outObj);
});

// WebSocket endpoint
app.UseWebSockets();

app.Map("/ws", async context =>
{
    if (!context.WebSockets.IsWebSocketRequest) { context.Response.StatusCode = 400; return; }
    var ws = await context.WebSockets.AcceptWebSocketAsync();
    sockets.Add(ws);
    Console.WriteLine("WebSocket connected. clients: " + sockets.Count);

    // keep the socket open; we do not expect the client to send, but we read to detect close
    var buffer = new byte[1024 * 4];
    try
    {
        while (ws.State == WebSocketState.Open)
        {
            var result = await ws.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
            if (result.CloseStatus.HasValue) break;
            // ignore application messages for now
            await Task.Delay(100);
        }
    }
    catch { /* ignore */ }
    finally
    {
        sockets.Remove(ws);
        try { await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "closing", CancellationToken.None); } catch { }
        Console.WriteLine("WebSocket disconnected. clients: " + sockets.Count);
    }
});

app.Run();

// Allow WebApplicationFactory<T> discovery in tests
public partial class Program { }

public record ToolInvokeRequest(string name, Dictionary<string, string>? args);

// Use helpers from Core

// Request DTOs for tool invocations
public record ListTablesRequest(string sheetName);
public record GetSheetDataRequest(string sheetName);
public record GetTableDataRequest(string sheetName, string tableName);

// For WebApplicationFactory in tests
public partial class Program { }