# Excel MCP Distribution Guide

This guide walks through packaging the Excel MCP tools, sharing them with end users, and running them in a few common scenarios (self-service chat UI, standalone server, and headless integration).

## 1. Build the Bundles

Run the packaging script from the repository root. Choose one or more Runtime Identifiers (RIDs) that match the target machines.

```powershell
pwsh -File "scripts/package.ps1" -Runtime @('win-x64','linux-x64')
```

Key outputs:

- `dist/<rid>/ExcelMcp.Server/ExcelMcp.Server[.exe]`
- `dist/<rid>/ExcelMcp.Client/ExcelMcp.Client[.exe]`
- `dist/<rid>/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb[.exe]`
- Launcher scripts: `run-client`, `run-server`, `run-chat` (`.ps1`, `.bat`, `.sh`)
- Optional archives: `excel-mcp-<rid>.zip` (omit with `-SkipZip` if you only need the folders)

> **Tip:** The publish step creates self-contained executables; the .NET runtime is bundled and no additional installs are required.

## 2. Share with Users

Distribute either of the following:

- The zip file produced in `dist` (easiest for email or download portals), **or**
- The entire runtime folder `dist/<rid>` copied onto removable media or a shared drive.

Instruct the user to extract the archive (or copy the folder) to a writable location, e.g.:

```
C:\Apps\ExcelMcp\win-x64
~/excel-mcp/linux-x64
```

All executables and static assets (including `wwwroot` for the chat UI) are packaged within that directory tree.

## 3. Use Cases and Examples

### 3.1 Launch the Conversational Web UI (Windows)

```powershell
cd "C:\Apps\ExcelMcp\win-x64"
./run-chat.ps1
```

- The script prompts for a workbook path (for example `D:\Data\finance.xlsx`).
- It locates the bundled server, sets the required environment variables, and launches `ExcelMcp.ChatWeb.exe`.
- Browse to the URL shown (defaults to `http://localhost:5000`).

For quick launch via double-click, use `run-chat.bat` in the same folder; it calls the PowerShell script behind the scenes.

### 3.2 Launch the Conversational Web UI (Linux / macOS)

```bash
cd ~/excel-mcp/linux-x64
chmod +x run-chat.sh
./run-chat.sh --workbook /home/user/data/finance.xlsx --urls http://0.0.0.0:8080
```

The script exports `EXCEL_MCP_WORKBOOK` and points to the bundled server before starting the chat web app. Adjust `--urls` to bind to a different host/port if needed.

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

- The process hosts the MCP tools over standard input/output.
- If `--workbook` is omitted and the console is interactive, the executable prompts for the path.
- Terminate with `Ctrl+C` when finished.

### 3.4 Integrate with Your Own Chat Completion Agent

When embedding into another application:

1. Ensure `ExcelMcp.Server` is on disk next to your agent.
2. Spawn the process with the workbook argument and capture stdio pipes.
3. Speak MCP JSON-RPC over those pipes to list tools or invoke them.

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
// Connect the pipes to your MCP transport handler
```

You may reuse the packaged `ExcelMcp.Client` to explore available tools from the command line:

```powershell
./run-client.ps1 -WorkbookPath "D:\Data\finance.xlsx" list
```

```bash
./run-client.sh --workbook /home/user/data/finance.xlsx search --query "Contoso"
```

## 4. Troubleshooting

- **Static file warning for ChatWeb**: Ensure the user extracted the full `ExcelMcp.ChatWeb` folder; the `wwwroot` directory must sit beside the executable.
- **Workbook prompt repeats**: The scripts set `EXCEL_MCP_WORKBOOK` for the lifetime of the process. Provide the `--workbook` argument or export the variable ahead of time to avoid prompts in automated scenarios.
- **MCP connection errors**: Verify the workbook path exists and is readable. The server logs to stderr; capture it when embedding in your own agent to diagnose issues.

## 5. Regenerating Bundles After Changes

Whenever code changes are made:

1. Run `dotnet test` (optional but recommended).
2. Re-run the packaging script for each target runtime.
3. Distribute the updated zip/folder.

```powershell
pwsh -File "scripts/package.ps1" -Runtime win-x64
```

Add `-SkipZip` to speed up local testing, or include multiple RIDs in one invocation.

---
Questions or feedback? Open an issue or reach out via the project’s discussion channels.
