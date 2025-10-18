# Excel MCP Distribution Guide

This guide walks through packaging the Excel MCP apps, sharing them with end users, and running them in common scenarios (chat UI, standalone server, and integrations).

## 1. Build the Bundles

Each component now has its own packaging script. From the repository root run whichever bundles you need, optionally passing one or more Runtime Identifiers (RIDs) such as `win-x64`, `linux-x64`, or `linux-arm64`.

```powershell
pwsh -File scripts/package-server.ps1
pwsh -File scripts/package-client.ps1 -Runtime @('win-x64','linux-x64')
pwsh -File scripts/package-chatweb.ps1 -Runtime linux-x64 -SkipZip
```

Each script publishes a self-contained single-file executable, copies launch helpers (`run-*.ps1`, `run-*.bat`, `run-*.sh`), writes a README, and optionally creates a zip archive.

Key outputs per RID:

- `dist/<rid>/ExcelMcp.Server/ExcelMcp.Server[.exe]`
- `dist/<rid>/ExcelMcp.Client/ExcelMcp.Client[.exe]`
- `dist/<rid>/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb[.exe]` plus `wwwroot/`

> **Tip:** The .NET runtime is bundled inside the published executable. Users do not need any additional installs.

### LM Studio (or OpenAI-Compatible Server) Setup

The chat front end requires an OpenAI-style local inference server to handle `/v1/chat/completions` requests. LM Studio is the default, but any compatible service (Azure AI Foundry local endpoint, open-source bridges, etc.) works as long as it exposes the same REST contract.

Suggested LM Studio steps:

1. Install [LM Studio](https://lmstudio.ai/) and start the desktop app.
2. Open **Local Server**, download `phi-4-mini-reasoning`, and load it.
3. Confirm the server is running on `http://localhost:1234` (or note the port you choose) with the Serve API enabled.
4. Leave LM Studio running while you use the chat UI.

Using another OpenAI-compatible host?

- Point `src/ExcelMcp.ChatWeb/appsettings.json` (or the packaged copy under `dist/<rid>/ExcelMcp.ChatWeb`) to the new endpoint by editing:

    ```json
    "LlmStudio": {
        "BaseUrl": "http://localhost:5000",
        "Model": "my-other-model"
    }
    ```

- Restart `ExcelMcp.ChatWeb` after making changes. Provide any required API keys via environment variables if your alternative server enforces authentication.

## 2. Share with Users

Distribute either of the following:

- The zip file the script produced (`excel-mcp-<app>-<rid>.zip`), **or**
- The entire folder `dist/<rid>/<AppName>` copied to removable media or a shared drive.

Instruct users to extract or copy to a writable location, for example:

```
C:\Apps\ExcelMcp\win-x64
~/excel-mcp/linux-x64
```

Everything the app needs—including the ChatWeb `wwwroot` assets—is contained within that directory tree.

## 3. Use Cases and Examples

### 3.1 Launch the Chat Web UI (Windows)

```powershell
cd "C:\Apps\ExcelMcp\win-x64\ExcelMcp.ChatWeb"
./run-chatweb.ps1
```

- The launcher prompts for a workbook path (e.g., `D:\Data\finance.xlsx`).
- It locates the bundled server, sets environment variables, and starts `ExcelMcp.ChatWeb.exe`.
- Browse to the URL printed in the console (defaults to `http://localhost:5000`).

Double-click `run-chatweb.bat` for a shortcut that wraps the PowerShell script.

### 3.2 Launch the Chat Web UI (Linux / macOS)

```bash
cd ~/excel-mcp/linux-x64/ExcelMcp.ChatWeb
chmod +x run-chatweb.sh
./run-chatweb.sh --workbook /home/user/data/finance.xlsx --urls http://0.0.0.0:8080
```

The script exports `EXCEL_MCP_WORKBOOK`, points to the bundled server, and launches the ASP.NET app. Adjust `--urls` to bind to a different host/port.

### 3.3 Run the Server Manually (Any Platform)

Copy only the server folder onto a machine and start it with a workbook:

```powershell
cd "C:\Apps\ExcelMcp\win-x64\ExcelMcp.Server"
./ExcelMcp.Server.exe --workbook "D:\Data\finance.xlsx"
```

```bash
cd ~/excel-mcp/linux-x64/ExcelMcp.Server
./ExcelMcp.Server --workbook /home/user/data/finance.xlsx
```

- The process exposes MCP tools over standard input/output.
- If `--workbook` is omitted and the console is interactive, the executable prompts for a path.
- Stop with `Ctrl+C`.

### 3.4 Explore Tools with the CLI Client

```powershell
cd "C:\Apps\ExcelMcp\win-x64\ExcelMcp.Client"
./run-client.ps1 -WorkbookPath "D:\Data\finance.xlsx" list
```

```bash
cd ~/excel-mcp/linux-x64/ExcelMcp.Client
./run-client.sh --workbook /home/user/data/finance.xlsx search --query "Contoso"
```

The client wrapper prompts for paths when needed, sets environment variables, and forwards all remaining arguments to the executable.

### 3.5 Embed the Server in Your Own Agent

When launching the MCP server from another application:

1. Ensure `ExcelMcp.Server[.exe]` is alongside your agent binaries.
2. Spawn it with a workbook argument and capture stdio pipes.
3. Exchange MCP JSON-RPC messages over those pipes.

Example pseudo-code (C#):

```csharp
var start = new ProcessStartInfo
{
    FileName = "ExcelMcp.Server.exe",
    ArgumentList = { "--workbook", workbookPath },
    RedirectStandardInput = true,
    RedirectStandardOutput = true,
    RedirectStandardError = true,
    UseShellExecute = false
};
var process = Process.Start(start);
// Wire the streams into your MCP transport handler
```

### 3.6 Page Through Large Searches

`excel-search` returns up to 100 rows per call. When additional matches exist the JSON payload includes `"hasMore": true` and a `nextCursor` token. Supply that token on the next call to retrieve the following page.

- **CLI example**

    ```powershell
    dotnet run --project src/ExcelMcp.Client -- search --query "Dividend" --limit 25
    # Copy "nextCursor" from the JSON response, then:
    dotnet run --project src/ExcelMcp.Client -- search --query "Dividend" --limit 25 --cursor 25
    ```

- **Chat web** – the assistant surfaces `Next cursor: <value>` in the tool output and nudges the model to call the tool again with `{"cursor":"<value>"}` when more rows are available.

### 3.7 Preview Worksheets with Pagination

`excel-preview-table` supports the same cursor pattern so you can stream worksheets or tables in manageable pages. The tool now returns structured JSON plus CSV text. Use the JSON for automation and reference the CSV when you want a quick copy/paste view.

- **CLI example**

    ```powershell
    dotnet run --project src/ExcelMcp.Client -- preview --worksheet "Transactions" --rows 15
    # When the JSON block reports "nextCursor": "15":
    dotnet run --project src/ExcelMcp.Client -- preview --worksheet "Transactions" --rows 15 --cursor 15
    ```

    The CLI prints the structured JSON (headers, row numbers, `hasMore`, `nextCursor`) before the CSV snippet. Copy the cursor value when more pages remain.

- **Chat web Preview panel** – The panel beneath the chat input labelled *Preview Workbook Data* calls `excel-preview-table` directly. Enter the worksheet/table (worksheet names auto-complete from discovered resources), click **Preview**, and use **Load more** to fetch the next page without leaving the UI. Collapse it with **Hide** if you need more vertical space; the offset and next cursor remain visible when expanded again.

- **Agents** – When invoking the tool programmatically, supply `{"cursor":"<token>"}` alongside the original arguments to continue where you left off.

## 4. Troubleshooting

- **Missing static files for ChatWeb** – Confirm the `wwwroot` folder sits beside the executable. The packaging script copies it automatically; re-run the script if it is absent.
- **Workbook prompt repeats** – Provide `--workbook` to the launcher or set `EXCEL_MCP_WORKBOOK` before invoking the executable to avoid interactive prompts in automation.
- **MCP connection issues** – Ensure the workbook path exists and is readable. Capture stderr from the server process when embedding so you can review diagnostics.

## 5. Refreshing Bundles After Changes

Whenever you update the source:

1. Run `dotnet test` (recommended).
2. Re-run the relevant packaging scripts for each RID you ship.
3. Distribute the newly generated zip/folder.

```powershell
pwsh -File scripts/package-client.ps1 -Runtime win-x64
pwsh -File scripts/package-server.ps1 -SkipZip
```

Add `-SkipZip` when you only need the staged folder, or supply multiple RIDs in a single invocation.

---
Questions or feedback? Open an issue or reach out via the project’s discussion channels.
