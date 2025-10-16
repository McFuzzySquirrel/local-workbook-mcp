<#
.SYNOPSIS
    Builds self-contained distributable bundles for the Excel MCP server and client.

.DESCRIPTION
    Publishes the server and client projects as single-file, self-contained executables
    and assembles them (plus helper launch scripts and documentation) under dist/<runtime>.
    Optionally produces a zip archive for distribution.

.EXAMPLE
    pwsh -File scripts/package.ps1

.EXAMPLE
    pwsh -File scripts/package.ps1 -Runtime win-arm64 -SkipZip
#>
[CmdletBinding()]
param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release",
    [switch]$SkipZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$distRoot = Join-Path $repoRoot "dist"
$runtimeRoot = Join-Path $distRoot $Runtime
$serverTarget = Join-Path $runtimeRoot "ExcelMcp.Server"
$clientTarget = Join-Path $runtimeRoot "ExcelMcp.Client"

Write-Host "Cleaning output at $runtimeRoot" -ForegroundColor Cyan
if (Test-Path $runtimeRoot) {
    Remove-Item $runtimeRoot -Recurse -Force
}
New-Item -ItemType Directory -Path $serverTarget | Out-Null
New-Item -ItemType Directory -Path $clientTarget | Out-Null

$publishArgs = @(
    "-c", $Configuration,
    "-r", $Runtime,
    "--self-contained", "true",
    "/p:PublishSingleFile=true",
    "/p:IncludeNativeLibrariesForSelfExtract=true",
    "/p:EnableCompressionInSingleFile=true",
    "/p:PublishTrimmed=false"
)

Write-Host "Publishing ExcelMcp.Server ($Runtime)" -ForegroundColor Cyan
& dotnet publish (Join-Path $repoRoot "src/ExcelMcp.Server/ExcelMcp.Server.csproj") @publishArgs -o $serverTarget

Write-Host "Publishing ExcelMcp.Client ($Runtime)" -ForegroundColor Cyan
& dotnet publish (Join-Path $repoRoot "src/ExcelMcp.Client/ExcelMcp.Client.csproj") @publishArgs -o $clientTarget

$readmePath = Join-Path $runtimeRoot "README.txt"
$readmeContent = @"
Excel Local MCP Self-Contained Bundle
====================================

Contents
--------
- ExcelMcp.Server\ExcelMcp.Server.exe — MCP server (self-contained)
- ExcelMcp.Client\ExcelMcp.Client.exe — Command-line client
- run-client.ps1 / run-client.bat — Prompt for a workbook and run client commands
- run-server.ps1 — Launch the server and keep it running for MCP-aware tools

Quick Start
-----------
1. Place your workbook (for example, C:\Data\finance.xlsx) somewhere accessible.
2. Run .\run-client.ps1 and supply the workbook path when prompted, or call:
       .\run-client.ps1 -WorkbookPath "C:\Data\finance.xlsx" list
   Use other arguments like `search --query "Contoso"` or `preview --worksheet "Sales"`.
3. To run the server directly, call:
       .\run-server.ps1 -WorkbookPath "C:\Data\finance.xlsx"
   Stop it with Ctrl+C when finished or close the window.
4. For advanced automation, point MCP-aware applications at ExcelMcp.Server.exe with the
   `--workbook` argument.

Changing Workbooks
------------------
Re-run the scripts with a different -WorkbookPath value whenever you want to inspect
another workbook. Each invocation launches a fresh server process bound to the new file.

More Documentation
------------------
See the main project README for detailed CLI examples and OpenAI agent integration.
"@
Set-Content -Path $readmePath -Value $readmeContent -Encoding UTF8

$clientLauncherPs1 = @'
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ClientArgs
)

if (-not $WorkbookPath) {
    $WorkbookPath = Read-Host "Enter the full path to the Excel workbook"
}

if (-not (Test-Path $WorkbookPath)) {
    Write-Error "Workbook not found: $WorkbookPath"
    exit 1
}

$clientExe = Join-Path $PSScriptRoot "ExcelMcp.Client/ExcelMcp.Client.exe"
$serverExe = Join-Path $PSScriptRoot "ExcelMcp.Server/ExcelMcp.Server.exe"

if (-not (Test-Path $clientExe)) {
    Write-Error "Client executable not found at $clientExe"
    exit 1
}

if (-not (Test-Path $serverExe)) {
    Write-Error "Server executable not found at $serverExe"
    exit 1
}

if (-not $ClientArgs) {
    $ClientArgs = @()
}

$hasServerArg = $false
for ($i = 0; $i -lt $ClientArgs.Count; $i++) {
    if ($ClientArgs[$i] -eq "--server") {
        $hasServerArg = $true
        break
    }
}

if (-not $hasServerArg -and -not $env:EXCEL_MCP_SERVER) {
    $ClientArgs = @("--server", $serverExe) + $ClientArgs
}

& $clientExe --workbook $WorkbookPath @ClientArgs
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
'@
Set-Content -Path (Join-Path $runtimeRoot "run-client.ps1") -Value $clientLauncherPs1 -Encoding UTF8

$clientLauncherBat = '@echo off`r`npwsh -ExecutionPolicy Bypass -File "%~dp0run-client.ps1" %*`r`n'
Set-Content -Path (Join-Path $runtimeRoot "run-client.bat") -Value $clientLauncherBat -Encoding ASCII

$serverLauncherPs1 = @'
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath
)

if (-not $WorkbookPath) {
    $WorkbookPath = Read-Host "Enter the full path to the Excel workbook"
}

if (-not (Test-Path $WorkbookPath)) {
    Write-Error "Workbook not found: $WorkbookPath"
    exit 1
}

$serverExe = Join-Path $PSScriptRoot "ExcelMcp.Server/ExcelMcp.Server.exe"

if (-not (Test-Path $serverExe)) {
    Write-Error "Server executable not found at $serverExe"
    exit 1
}

Write-Host "Starting Excel MCP server for $WorkbookPath" -ForegroundColor Cyan
& $serverExe --workbook $WorkbookPath
'@
Set-Content -Path (Join-Path $runtimeRoot "run-server.ps1") -Value $serverLauncherPs1 -Encoding UTF8

if (-not $SkipZip) {
    $zipName = "excel-mcp-$Runtime.zip"
    $zipPath = Join-Path $distRoot $zipName
    Write-Host "Creating archive $zipPath" -ForegroundColor Cyan
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }
    Compress-Archive -Path (Join-Path $runtimeRoot '*') -DestinationPath $zipPath
}

Write-Host "Package complete: $runtimeRoot" -ForegroundColor Green
if (-not $SkipZip) {
    Write-Host "Archive created: $(Join-Path $distRoot $zipName)" -ForegroundColor Green
}
