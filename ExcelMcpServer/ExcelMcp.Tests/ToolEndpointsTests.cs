using System;
using System.IO;
using System.Net.Http.Json;
using Microsoft.AspNetCore.Mvc.Testing;
using Xunit;
using System.Threading.Tasks;
using System.Text.Json;

namespace ExcelMcp.Tests
{
    public class ToolEndpointsTests : IClassFixture<WebApplicationFactory<Program>>
    {
        private readonly WebApplicationFactory<Program> _factory;

        public ToolEndpointsTests(WebApplicationFactory<Program> factory)
        {
            _factory = factory.WithWebHostBuilder(builder =>
            {
                // Set a non-existing file to avoid reading during startup; endpoints load on demand
                builder.UseSetting("ExcelFile", Path.Combine(Path.GetTempPath(), Guid.NewGuid()+".xlsx"));
            });
        }

    [Fact]
    public async Task ListsTools()
        {
            var client = _factory.CreateClient();
            var resp = await client.GetAsync("/mcp/tools");
            resp.EnsureSuccessStatusCode();
            var json = await resp.Content.ReadAsStringAsync();
            Assert.Contains("listSheets", json);
        }

        [Fact]
        public async Task Invoke_ListSheets_ReturnsArray()
        {
            var client = _factory.CreateClient();
            var resp = await client.PostAsJsonAsync("/mcp/tools/invoke", new { name = "listSheets" });
            resp.EnsureSuccessStatusCode();
            var json = await resp.Content.ReadAsStringAsync();
            Assert.Contains("sheets", json);
        }

        [Fact]
        public async Task WriteTools_SetCell_And_AppendRow_Work()
        {
            // Create a temp workbook with a Sheet1
            var tmp = Path.Combine(Path.GetTempPath(), $"excelmcp_{Guid.NewGuid():N}.xlsx");
            using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.ReadWrite, FileShare.Read))
            {
                var wb = new NPOI.XSSF.UserModel.XSSFWorkbook();
                var sh = wb.CreateSheet("Sheet1");
                wb.Write(fs);
            }

            var factory = _factory.WithWebHostBuilder(builder =>
            {
                builder.UseSetting("ExcelFile", tmp);
            });
            var client = factory.CreateClient();

            // setCell B2 = Hello
            var setResp = await client.PostAsJsonAsync("/mcp/tools/invoke", new
            {
                name = "setCell",
                args = new { sheetName = "Sheet1", address = "B2", value = "Hello" }
            });
            if (!setResp.IsSuccessStatusCode)
            {
                var body = await setResp.Content.ReadAsStringAsync();
                throw new Exception($"setCell failed: {(int)setResp.StatusCode} {setResp.StatusCode} - {body}");
            }

            // appendRow values: X,Y,Z
            var appendResp = await client.PostAsJsonAsync("/mcp/tools/invoke", new
            {
                name = "appendRow",
                args = new { sheetName = "Sheet1", values = "X,Y,Z" }
            });
            if (!appendResp.IsSuccessStatusCode)
            {
                var body = await appendResp.Content.ReadAsStringAsync();
                throw new Exception($"appendRow failed: {(int)appendResp.StatusCode} {appendResp.StatusCode} - {body}");
            }

            // read back sheet data and assert contains Hello and X
            var readResp = await client.PostAsJsonAsync("/mcp/tools/invoke", new
            {
                name = "getSheetData",
                args = new { sheetName = "Sheet1" }
            });
            readResp.EnsureSuccessStatusCode();
            var payload = await readResp.Content.ReadFromJsonAsync<JsonElement>();
            Assert.True(payload.TryGetProperty("data", out var dataNode));
            var allText = dataNode.ToString();
            Assert.Contains("Hello", allText);
            Assert.Contains("X", allText);

            try { File.Delete(tmp); } catch { }
        }
    }
}
