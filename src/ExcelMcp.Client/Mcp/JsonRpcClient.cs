using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelMcp.Client.Mcp;

internal sealed class JsonRpcClient : IAsyncDisposable
{
    private readonly Stream _input;
    private readonly Stream _output;
    private readonly Encoding _encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
    private int _nextId;

    public JsonRpcClient(Stream input, Stream output)
    {
        _input = input;
        _output = output;
    }

    public async Task<JsonDocument> SendRequestAsync(string method, JsonNode? parameters, CancellationToken cancellationToken)
    {
        var id = Interlocked.Increment(ref _nextId);
        await WriteRequestAsync(id, method, parameters, cancellationToken).ConfigureAwait(false);
        return await ReadResponseAsync(id, cancellationToken).ConfigureAwait(false);
    }

    public async Task SendNotificationAsync(string method, JsonNode? parameters, CancellationToken cancellationToken)
    {
        await WriteMessageAsync(null, method, parameters, cancellationToken).ConfigureAwait(false);
    }

    private async Task WriteRequestAsync(int id, string method, JsonNode? parameters, CancellationToken cancellationToken)
    {
        await WriteMessageAsync(id, method, parameters, cancellationToken).ConfigureAwait(false);
    }

    private async Task WriteMessageAsync(int? id, string method, JsonNode? parameters, CancellationToken cancellationToken)
    {
        using var payload = new MemoryStream();
        using (var writer = new Utf8JsonWriter(payload))
        {
            writer.WriteStartObject();
            writer.WriteString("jsonrpc", "2.0");
            if (id is { } identifier)
            {
                writer.WriteNumber("id", identifier);
            }

            writer.WriteString("method", method);
            if (parameters is not null)
            {
                writer.WritePropertyName("params");
                parameters.WriteTo(writer);
            }

            writer.WriteEndObject();
        }

        await payload.FlushAsync(cancellationToken).ConfigureAwait(false);
        payload.Position = 0;
        var body = payload.ToArray();
        await _output.WriteAsync(body, cancellationToken).ConfigureAwait(false);
        await _output.WriteAsync(_encoding.GetBytes("\n"), cancellationToken).ConfigureAwait(false);
        await _output.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<JsonDocument> ReadResponseAsync(int expectedId, CancellationToken cancellationToken)
    {
        while (true)
        {
            var message = await ReadMessageAsync(cancellationToken).ConfigureAwait(false);
            if (message is null)
            {
                throw new EndOfStreamException("Stream closed before response was received.");
            }

            using (message)
            {
                if (message.RootElement.TryGetProperty("id", out var idElement) && idElement.ValueKind == JsonValueKind.Number && idElement.GetInt32() == expectedId)
                {
                    if (message.RootElement.TryGetProperty("error", out var errorElement))
                    {
                        var code = errorElement.TryGetProperty("code", out var codeElement) ? codeElement.GetInt32() : -32603;
                        var messageText = errorElement.TryGetProperty("message", out var messageElement) ? messageElement.GetString() : "Unknown error";
                        throw new InvalidOperationException($"JSON-RPC error {code}: {messageText}");
                    }

                    return JsonDocument.Parse(message.RootElement.GetRawText());
                }
            }
        }
    }

    private async Task<JsonDocument?> ReadMessageAsync(CancellationToken cancellationToken)
    {
        while (true)
        {
            var firstLine = await ReadLineAsync(cancellationToken).ConfigureAwait(false);
            if (firstLine is null)
            {
                return null;
            }

            if (firstLine.Length == 0)
            {
                continue;
            }

            var trimmed = firstLine.TrimStart();
            if (trimmed.StartsWith('{') || trimmed.StartsWith('['))
            {
                return JsonDocument.Parse(trimmed);
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

                    throw new EndOfStreamException("Unexpected end of stream while reading headers.");
                }

                if (line.Length == 0)
                {
                    break;
                }

                ProcessHeaderLine(line, headers);
            }

            if (!headers.TryGetValue("Content-Length", out var lengthValue) || !int.TryParse(lengthValue, out var length))
            {
                throw new InvalidOperationException("Missing Content-Length header.");
            }

            var buffer = new byte[length];
            var read = 0;
            while (read < length)
            {
                var bytesRead = await _input.ReadAsync(buffer.AsMemory(read, length - read), cancellationToken).ConfigureAwait(false);
                if (bytesRead == 0)
                {
                    throw new EndOfStreamException("Unexpected end of stream while reading JSON-RPC response.");
                }

                read += bytesRead;
            }

            return JsonDocument.Parse(buffer);
        }
    }

    private async Task<string?> ReadLineAsync(CancellationToken cancellationToken)
    {
        var buffer = new List<byte>();
        while (true)
        {
            var single = new byte[1];
            var read = await _input.ReadAsync(single.AsMemory(0, 1), cancellationToken).ConfigureAwait(false);
            if (read == 0)
            {
                return buffer.Count == 0 ? null : throw new EndOfStreamException("Unexpected end of stream while reading header line.");
            }

            var current = single[0];
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

    private static void ProcessHeaderLine(string line, Dictionary<string, string> headers)
    {
        var separatorIndex = line.IndexOf(':');
        if (separatorIndex <= 0)
        {
            return;
        }

        var name = line[..separatorIndex].Trim();
        var value = line[(separatorIndex + 1)..].Trim();
        headers[name] = value;
    }

    public ValueTask DisposeAsync()
    {
        return ValueTask.CompletedTask;
    }
}
