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
        var headers = await ReadHeadersAsync(cancellationToken).ConfigureAwait(false);
        if (headers is null)
        {
            return null;
        }

        if (!headers.TryGetValue("Content-Length", out var lengthValue) || !int.TryParse(lengthValue, out var length))
        {
            throw new InvalidOperationException("Missing Content-Length header in JSON-RPC message.");
        }

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
    return new JsonRpcMessage(document.RootElement.Clone());
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
        var length = (int)payload.Length;
        var header = _encoding.GetBytes($"Content-Length: {length}\r\nContent-Type: application/json\r\n\r\n");

        await _output.WriteAsync(header, cancellationToken).ConfigureAwait(false);
        payload.Position = 0;
        await payload.CopyToAsync(_output, cancellationToken).ConfigureAwait(false);
        await _output.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<Dictionary<string, string>?> ReadHeadersAsync(CancellationToken cancellationToken)
    {
    var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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
                return headers;
            }

            var separatorIndex = line.IndexOf(':');
            if (separatorIndex <= 0)
            {
                continue;
            }

            var name = line[..separatorIndex].Trim();
            var value = line[(separatorIndex + 1)..].Trim();
            headers[name] = value;
        }
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
}
