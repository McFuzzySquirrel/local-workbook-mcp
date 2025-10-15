using System.Net.Http.Json;
using System.Net.WebSockets;
using System.Text;

var baseUrl = args.FirstOrDefault() ?? "http://localhost:5089";
Console.WriteLine($"MCP Excel client targeting {baseUrl}");

using var http = new HttpClient { BaseAddress = new Uri(baseUrl) };

// list tools
var tools = await http.GetFromJsonAsync<object>("/mcp/tools");
Console.WriteLine("/mcp/tools ->\n" + System.Text.Json.JsonSerializer.Serialize(tools));

// list resources
var resources = await http.GetFromJsonAsync<object>("/mcp/resources");
Console.WriteLine("/mcp/resources ->\n" + System.Text.Json.JsonSerializer.Serialize(resources));

// invoke listSheets
var result = await http.PostAsJsonAsync("/mcp/tools/invoke", new { name = "listSheets" });
Console.WriteLine("invoke listSheets ->\n" + await result.Content.ReadAsStringAsync());

// open WebSocket to receive change events
Console.WriteLine("Connecting WebSocket /ws (press Ctrl+C to exit)...");
using var ws = new ClientWebSocket();
await ws.ConnectAsync(new Uri(new Uri(baseUrl), "/ws"), CancellationToken.None);

var buffer = new byte[8192];
while (ws.State == WebSocketState.Open)
{
    var res = await ws.ReceiveAsync(buffer, CancellationToken.None);
    if (res.MessageType == WebSocketMessageType.Close) break;
    var text = Encoding.UTF8.GetString(buffer, 0, res.Count);
    Console.WriteLine($"Event: {text}");
}

Console.WriteLine("WebSocket closed.");
