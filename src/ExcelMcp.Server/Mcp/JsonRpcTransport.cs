using System.Text;
using System.Text.Json;

namespace ExcelMcp.Server.Mcp;

internal sealed class JsonRpcTransport : IAsyncDisposable
{
    private readonly Stream _input;
    private readonly Stream _output;
    private readonly Encoding _encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    public JsonRpcTransport(Stream input, Stream output)
    {
        _input = input;
        _output = output;
    }

    public async Task<JsonRpcMessage?> ReadAsync(CancellationToken cancellationToken)
    {
        while (true)
        {
            Log("Waiting for header or JSON line");
            var firstLine = await ReadLineAsync(cancellationToken).ConfigureAwait(false);
            if (firstLine is null)
            {
                Log("Stream ended before receiving data");
                return null;
            }

            if (firstLine.Length == 0)
            {
                Log("Skipping empty line");
                continue;
            }

            var trimmed = firstLine.TrimStart();
            if (trimmed.StartsWith('{') || trimmed.StartsWith('['))
            {
                Log($"Received raw JSON line: {trimmed}");
                using var inlineDocument = JsonDocument.Parse(trimmed);
                return new JsonRpcMessage(inlineDocument.RootElement.Clone());
            }

            var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            ProcessHeaderLine(firstLine, headers);

            while (true)
            {
                var line = await ReadLineAsync(cancellationToken).ConfigureAwait(false);
                if (line is null)
                {
                    if (headers.Count == 0)
                    {
                        return null;
                    }

                    throw new EndOfStreamException("Unexpected end of stream while reading JSON-RPC headers.");
                }

                if (line.Length == 0)
                {
                    Log("Reached end of headers");
                    break;
                }

                ProcessHeaderLine(line, headers);
            }

            if (!headers.TryGetValue("Content-Length", out var lengthValue) || !int.TryParse(lengthValue, out var length))
            {
                throw new InvalidOperationException("Missing Content-Length header in JSON-RPC message.");
            }

            Log($"Content-Length: {length}");
            var buffer = new byte[length];
            var read = 0;
            while (read < length)
            {
                var bytesRead = await _input.ReadAsync(buffer.AsMemory(read, length - read), cancellationToken).ConfigureAwait(false);
                if (bytesRead == 0)
                {
                    throw new EndOfStreamException("Unexpected end of stream while reading JSON-RPC body.");
                }

                read += bytesRead;
            }

            using var document = JsonDocument.Parse(buffer);
            Log($"Read JSON message: {document.RootElement.GetRawText()}");
            return new JsonRpcMessage(document.RootElement.Clone());
        }
    }

    public async Task WriteResultAsync(JsonElement? id, object result, CancellationToken cancellationToken)
    {
        using var memory = new MemoryStream();
        using (var writer = new Utf8JsonWriter(memory))
        {
            writer.WriteStartObject();
            writer.WriteString("jsonrpc", "2.0");
            if (id is { } identifier)
            {
                writer.WritePropertyName("id");
                identifier.WriteTo(writer);
            }

            writer.WritePropertyName("result");
            System.Text.Json.JsonSerializer.Serialize(writer, result, JsonOptions.Serializer);
            writer.WriteEndObject();
        }

        await WriteMessageAsync(memory, cancellationToken).ConfigureAwait(false);
    }

    public async Task WriteErrorAsync(JsonElement? id, McpErrorResponse error, CancellationToken cancellationToken)
    {
        using var memory = new MemoryStream();
        using (var writer = new Utf8JsonWriter(memory))
        {
            writer.WriteStartObject();
            writer.WriteString("jsonrpc", "2.0");
            if (id is { } identifier)
            {
                writer.WritePropertyName("id");
                identifier.WriteTo(writer);
            }

            writer.WritePropertyName("error");
            System.Text.Json.JsonSerializer.Serialize(writer, error, JsonOptions.Serializer);
            writer.WriteEndObject();
        }

        await WriteMessageAsync(memory, cancellationToken).ConfigureAwait(false);
    }

    private async Task WriteMessageAsync(MemoryStream payload, CancellationToken cancellationToken)
    {
        await payload.FlushAsync(cancellationToken).ConfigureAwait(false);
        payload.Position = 0;
        var body = payload.ToArray();
        Log($"Sending JSON message: {Encoding.UTF8.GetString(body)}");
        await _output.WriteAsync(body, cancellationToken).ConfigureAwait(false);
        await _output.WriteAsync(_encoding.GetBytes("\n"), cancellationToken).ConfigureAwait(false);
        await _output.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<string?> ReadLineAsync(CancellationToken cancellationToken)
    {
        var buffer = new List<byte>();
        while (true)
        {
            var temp = new byte[1];
            var read = await _input.ReadAsync(temp.AsMemory(0, 1), cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                return buffer.Count == 0 ? null : throw new EndOfStreamException("Unexpected end of stream while reading header line.");
            }

            var current = temp[0];
            if (current == '\r')
            {
                var next = new byte[1];
                var nextRead = await _input.ReadAsync(next.AsMemory(0, 1), cancellationToken).ConfigureAwait(false);
                if (nextRead == 0)
                {
                    throw new EndOfStreamException("Unexpected end of stream while reading CRLF.");
                }

                if (next[0] == '\n')
                {
                    break;
                }

                buffer.Add(current);
                buffer.Add(next[0]);
                continue;
            }

            if (current == '\n')
            {
                break;
            }

            buffer.Add(current);
        }

        return Encoding.ASCII.GetString(buffer.ToArray());
    }

    public ValueTask DisposeAsync()
    {
        return ValueTask.CompletedTask;
    }

    private static void ProcessHeaderLine(string line, Dictionary<string, string> headers)
    {
        var separatorIndex = line.IndexOf(':');
        if (separatorIndex <= 0)
        {
            Log($"Skipping malformed header line: {line}");
            return;
        }

        var name = line[..separatorIndex].Trim();
        var value = line[(separatorIndex + 1)..].Trim();
        headers[name] = value;
        Log($"Header '{name}': '{value}'");
    }

    private static void Log(string message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return;
        }

        try
        {
            Console.Error.WriteLine($"[{DateTimeOffset.UtcNow:O}] transport: {message}");
        }
        catch
        {
            // ignore logging failures
        }
    }
}
