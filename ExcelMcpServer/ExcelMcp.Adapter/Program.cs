using System.Buffers;
using System.Data;
using System.Text;
using System.Text.Json;
using ExcelMcp.Core;

string excelFile = Environment.GetEnvironmentVariable("ExcelFile")
                     ?? (args.Length > 0 ? args[0] : "data.xlsx");

ToolDefinition[] ToolDefinitions = new[]
{
    new ToolDefinition(
        "listSheets",
        "Return sheet names",
        """{ "type": "object", "properties": {}, "required": [] }"""),
    new ToolDefinition(
        "listTables",
        "Return table names for a sheet",
        """{ "type": "object", "properties": { "sheetName": { "type": "string" } }, "required": ["sheetName"] }"""),
    new ToolDefinition(
        "getSheetData",
        "Return raw cell matrix for a sheet",
        """{ "type": "object", "properties": { "sheetName": { "type": "string" } }, "required": ["sheetName"] }"""),
    new ToolDefinition(
        "getTableData",
        "Return rows for a specific table",
        """{ "type": "object", "properties": { "sheetName": { "type": "string" }, "tableName": { "type": "string" } }, "required": ["sheetName", "tableName"] }"""),
    new ToolDefinition(
        "setCell",
        "Set a single cell by A1 address",
        """{ "type": "object", "properties": { "sheetName": { "type": "string" }, "address": { "type": "string" }, "value": { "type": "string" } }, "required": ["sheetName", "address", "value"] }"""),
    new ToolDefinition(
        "appendRow",
        "Append a row to the end of the sheet",
        """{ "type": "object", "properties": { "sheetName": { "type": "string" }, "values": { "type": "array", "items": { "type": "string" } } }, "required": ["sheetName", "values"] }"""),
};

var rpc = new StdioRpc();
await rpc.RunAsync(HandleRequestAsync);

async Task HandleRequestAsync(JsonElement request)
{
    var id = request.TryGetProperty("id", out var idNode) ? idNode : default;
    var method = request.GetProperty("method").GetString();
    var @params = request.TryGetProperty("params", out var p) ? p : default;

    try
    {
        Console.Error.WriteLine($"[rpc] method={method}");
        switch (method)
        {
            case "initialize":
                await rpc.RespondAsync(id, static writer =>
                {
                    writer.WriteStartObject();
                    writer.WriteString("protocolVersion", "2024-11-05");
                    writer.WritePropertyName("serverInfo");
                    writer.WriteStartObject();
                    writer.WriteString("name", "excel-workbook-mcp");
                    writer.WriteString("version", "0.1.0");
                    writer.WriteEndObject();
                    writer.WritePropertyName("capabilities");
                    writer.WriteStartObject();
                    writer.WritePropertyName("tools");
                    writer.WriteStartObject();
                    writer.WriteBoolean("listChanged", false);
                    writer.WriteEndObject();
                    writer.WriteEndObject();
                    writer.WriteEndObject();
                });
                break;

            case "tools/list":
                await rpc.RespondAsync(id, writer =>
                {
                    writer.WriteStartObject();
                    writer.WritePropertyName("tools");
                    writer.WriteStartArray();
                    foreach (var tool in ToolDefinitions)
                    {
                        tool.WriteDescriptor(writer);
                    }
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                });
                break;

            case "tools/call":
                var name = @params.GetProperty("name").GetString()!;
                var arguments = @params.TryGetProperty("arguments", out var a) && a.ValueKind == JsonValueKind.Object ? a : default;
                var result = CallTool(name, arguments);
                await rpc.RespondAsync(id, writer => WriteToolCallResult(writer, result));
                break;

            case "shutdown":
                await rpc.RespondAsync(id, static writer => writer.WriteNullValue());
                break;

            case "exit":
                Environment.Exit(0);
                break;

            default:
                await rpc.ErrorAsync(id, -32601, $"Unknown method {method}");
                break;
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"[rpc-error] {ex}");
        await rpc.ErrorAsync(id, -32000, ex.Message);
    }
}

static void WriteToolCallResult(Utf8JsonWriter writer, ToolResult result)
{
    writer.WriteStartObject();
    writer.WritePropertyName("content");
    writer.WriteStartArray();
    writer.WriteStartObject();
    writer.WriteString("type", "json");
    writer.WritePropertyName("data");
    result.Write(writer);
    writer.WriteEndObject();
    writer.WriteEndArray();
    writer.WriteEndObject();
}

ToolResult CallTool(string name, JsonElement args)
{
    switch (name)
    {
        case "listSheets":
            {
                var data = ExcelReader.LoadWorkbook(excelFile);
                var sheets = data.Keys.ToArray();
                return new ToolResult(writer =>
                {
                    writer.WriteStartObject();
                    writer.WritePropertyName("sheets");
                    writer.WriteStartArray();
                    foreach (var sheet in sheets)
                    {
                        writer.WriteStringValue(sheet);
                    }
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                });
            }
        case "listTables":
            {
                var sheet = GetRequiredString(args, "sheetName");
                var data = ExcelReader.LoadWorkbook(excelFile);
                var tables = data.TryGetValue(sheet, out var list)
                    ? list.Select(t => t.TableName).ToArray()
                    : Array.Empty<string>();
                return new ToolResult(writer =>
                {
                    writer.WriteStartObject();
                    writer.WritePropertyName("tables");
                    writer.WriteStartArray();
                    foreach (var table in tables)
                    {
                        writer.WriteStringValue(table);
                    }
                    writer.WriteEndArray();
                    writer.WriteEndObject();
                });
            }
        case "getSheetData":
            {
                var sheet = GetRequiredString(args, "sheetName");
                var matrix = ExcelReadHelpers.ReadSheetAsMatrix(excelFile, sheet);
                return new ToolResult(writer => WriteMatrix(writer, matrix));
            }
        case "getTableData":
            {
                var sheet = GetRequiredString(args, "sheetName");
                var table = GetRequiredString(args, "tableName");
                var data = ExcelReader.LoadWorkbook(excelFile);
                var rows = Array.Empty<string[]>();
                if (data.TryGetValue(sheet, out var tables))
                {
                    var tbl = tables.FirstOrDefault(t => string.Equals(t.TableName, table, StringComparison.OrdinalIgnoreCase));
                    if (tbl != null)
                    {
                        rows = tbl.Rows.Cast<DataRow>()
                            .Select(r => r.ItemArray.Select(c => c?.ToString() ?? string.Empty).ToArray())
                            .ToArray();
                    }
                }
                var captured = rows;
                return new ToolResult(writer => WriteMatrix(writer, captured));
            }
        case "setCell":
            {
                var sheet = GetRequiredString(args, "sheetName");
                var address = GetRequiredString(args, "address");
                var value = args.TryGetProperty("value", out var v) ? (v.GetString() ?? string.Empty) : string.Empty;
                ExcelWriteHelpers.WithRetry(() => ExcelWriteHelpers.SetCellValue(excelFile, sheet, address, value));
                return OkResult();
            }
        case "appendRow":
            {
                var sheet = GetRequiredString(args, "sheetName");
                string[] values;
                if (args.TryGetProperty("values", out var v))
                {
                    if (v.ValueKind == JsonValueKind.Array)
                        values = v.EnumerateArray().Select(x => x.GetString() ?? string.Empty).ToArray();
                    else
                        values = (v.GetString() ?? string.Empty).Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                }
                else
                {
                    values = Array.Empty<string>();
                }
                ExcelWriteHelpers.WithRetry(() => ExcelWriteHelpers.AppendRow(excelFile, sheet, values));
                return OkResult();
            }
        default:
            throw new InvalidOperationException($"Unknown tool {name}");
    }
}

static ToolResult OkResult() => new(writer =>
{
    writer.WriteStartObject();
    writer.WriteBoolean("ok", true);
    writer.WriteEndObject();
});

static void WriteMatrix(Utf8JsonWriter writer, string[][] matrix)
{
    writer.WriteStartObject();
    writer.WritePropertyName("data");
    writer.WriteStartArray();
    foreach (var row in matrix)
    {
        writer.WriteStartArray();
        foreach (var cell in row)
        {
            writer.WriteStringValue(cell ?? string.Empty);
        }
        writer.WriteEndArray();
    }
    writer.WriteEndArray();
    writer.WriteEndObject();
}

static string GetRequiredString(JsonElement args, string propertyName)
{
    if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(propertyName, out var prop) || prop.ValueKind == JsonValueKind.Null)
    {
        throw new ArgumentException($"Argument '{propertyName}' is required.");
    }
    if (prop.ValueKind != JsonValueKind.String)
    {
        throw new ArgumentException($"Argument '{propertyName}' must be a string.");
    }
    var value = prop.GetString();
    if (string.IsNullOrWhiteSpace(value))
    {
        throw new ArgumentException($"Argument '{propertyName}' cannot be empty.");
    }
    return value;
}

readonly struct ToolResult
{
    private readonly Action<Utf8JsonWriter> _writer;

    public ToolResult(Action<Utf8JsonWriter> writer)
    {
        _writer = writer;
    }

    public void Write(Utf8JsonWriter writer) => _writer(writer);
}

sealed class ToolDefinition
{
    private readonly JsonDocument _schema;

    public ToolDefinition(string name, string description, string schemaJson)
    {
        Name = name;
        Description = description;
        _schema = JsonDocument.Parse(schemaJson);
    }

    public string Name { get; }
    public string Description { get; }

    public void WriteDescriptor(Utf8JsonWriter writer)
    {
        writer.WriteStartObject();
        writer.WriteString("name", Name);
        writer.WriteString("description", Description);
        writer.WritePropertyName("inputSchema");
        _schema.RootElement.WriteTo(writer);
        writer.WriteEndObject();
    }
}

sealed class StdioRpc
{
    private readonly Stream _in = Console.OpenStandardInput();
    private readonly Stream _out = Console.OpenStandardOutput();
    private readonly byte[] _headerBuf = new byte[8192];

    public async Task RunAsync(Func<JsonElement, Task> onRequest)
    {
        while (true)
        {
            var payload = await ReadMessageAsync();
            if (payload.ValueKind == JsonValueKind.Undefined) break;
            await onRequest(payload);
        }
    }

    public Task RespondAsync(JsonElement id, Action<Utf8JsonWriter> writeResult)
        => WriteResponseAsync(id, writeResult, null);

    public Task ErrorAsync(JsonElement id, int code, string message)
        => WriteResponseAsync(id, null, writer =>
        {
            writer.WriteStartObject();
            writer.WriteNumber("code", code);
            writer.WriteString("message", message);
            writer.WriteEndObject();
        });

    private async Task<JsonElement> ReadMessageAsync()
    {
        int total = 0;
        int contentLength = -1;
        while (true)
        {
            int read = await ReadLineAsync(_headerBuf, total);
            if (read <= 0) return default;
            var raw = Encoding.ASCII.GetString(_headerBuf, total, read);
            total += read;
            var line = raw.TrimEnd('\r', '\n');
            if (line.Length == 0) break;
            if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
            {
                var val = line.Substring("Content-Length:".Length).Trim();
                int.TryParse(val, out contentLength);
            }
            if (total >= _headerBuf.Length - 2) total = 0;
        }

        if (contentLength <= 0) return default;
        var payloadBuf = ArrayPool<byte>.Shared.Rent(contentLength);
        try
        {
            int offset = 0;
            while (offset < contentLength)
            {
                int n = await _in.ReadAsync(payloadBuf, offset, contentLength - offset);
                if (n <= 0) break;
                offset += n;
            }
            var doc = JsonDocument.Parse(new ReadOnlyMemory<byte>(payloadBuf, 0, contentLength));
            return doc.RootElement.Clone();
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(payloadBuf);
        }
    }

    private async Task WriteResponseAsync(JsonElement id, Action<Utf8JsonWriter>? writeResult, Action<Utf8JsonWriter>? writeError)
    {
        using var ms = new MemoryStream();
        using (var writer = new Utf8JsonWriter(ms))
        {
            writer.WriteStartObject();
            writer.WriteString("jsonrpc", "2.0");
            writer.WritePropertyName("id");
            if (id.ValueKind == JsonValueKind.Undefined)
            {
                writer.WriteNullValue();
            }
            else
            {
                id.WriteTo(writer);
            }

            if (writeResult is not null)
            {
                writer.WritePropertyName("result");
                writeResult(writer);
            }
            else if (writeError is not null)
            {
                writer.WritePropertyName("error");
                writeError(writer);
            }

            writer.WriteEndObject();
            writer.Flush();
        }

        var bytes = ms.ToArray();
        var header = Encoding.ASCII.GetBytes($"Content-Length: {bytes.Length}\r\n\r\n");
        await _out.WriteAsync(header, 0, header.Length);
        await _out.WriteAsync(bytes, 0, bytes.Length);
        await _out.FlushAsync();
    }

    private async Task<int> ReadLineAsync(byte[] buf, int start)
    {
        int i = start;
        while (true)
        {
            int b = _in.ReadByte();
            if (b == -1) return -1;
            buf[i++] = (byte)b;
            if (buf[i - 1] == (byte)'\n')
                return i - start;
            if (i >= buf.Length) return -1;
            await Task.Yield();
        }
    }
}
