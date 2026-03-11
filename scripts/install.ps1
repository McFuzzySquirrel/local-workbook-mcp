# install.ps1 — Build and install ExcelMcp.Server and ExcelMcp.ChatWeb on Windows
param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

$RepoRoot   = (Resolve-Path "$PSScriptRoot/..").Path
$PublishDir = Join-Path $RepoRoot "publish"
$ServerOut  = Join-Path $PublishDir "server"
$ChatWebOut = Join-Path $PublishDir "chatweb"

Write-Host "=== local-workbook-mcp install ===" -ForegroundColor Cyan
Write-Host "Repo root : $RepoRoot"
Write-Host "Output    : $PublishDir"
Write-Host ""

# Publish MCP server
Write-Host "[1/2] Publishing ExcelMcp.Server..." -ForegroundColor Green
dotnet publish "$RepoRoot/src/ExcelMcp.Server/ExcelMcp.Server.csproj" `
    --configuration $Configuration `
    --output $ServerOut `
    --nologo

# Publish Blazor ChatWeb
Write-Host "[2/2] Publishing ExcelMcp.ChatWeb..." -ForegroundColor Green
dotnet publish "$RepoRoot/src/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb.csproj" `
    --configuration $Configuration `
    --output $ChatWebOut `
    --nologo

Write-Host ""
Write-Host "Done! Binaries written to:" -ForegroundColor Cyan
Write-Host "  Server  : $ServerOut\ExcelMcp.Server.exe"
Write-Host "  ChatWeb : $ChatWebOut\ExcelMcp.ChatWeb.exe"
Write-Host ""
Write-Host "Quick-start:" -ForegroundColor Yellow
Write-Host "  # Run MCP server standalone (for Claude Desktop / Cursor / VS Code)"
Write-Host "  & `"$ServerOut\ExcelMcp.Server.exe`" --workbook C:\path\to\workbook.xlsx"
Write-Host ""
Write-Host "  # Run Blazor Chat UI"
Write-Host "  `$env:EXCEL_MCP_SERVER=`"$ServerOut\ExcelMcp.Server.exe`""
Write-Host "  & `"$ChatWebOut\ExcelMcp.ChatWeb.exe`""
Write-Host ""
Write-Host "See mcp-config/ for client config templates."
