<#
.SYNOPSIS
    Publishes the Excel MCP server as a self-contained, single-file executable bundle.

.DESCRIPTION
    For each requested Runtime Identifier (RID) the script performs a self-contained
    publish of ExcelMcp.Server, stages the output under dist/<rid>/ExcelMcp.Server,
    and adds small launch helpers so end users can provide workbook paths interactively.
    Optionally, a zip archive is produced per RID for easy distribution.

.EXAMPLE
    pwsh -File scripts/package-server.ps1

.EXAMPLE
    pwsh -File scripts/package-server.ps1 -Runtime @('win-x64','linux-x64')

.EXAMPLE
    pwsh -File scripts/package-server.ps1 -Runtime win-x64 -SkipZip
#>
[CmdletBinding()]
param(
    [string[]]$Runtime = @('win-x64'),
    [string]$Configuration = 'Release',
    [switch]$SkipZip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$distRoot = Join-Path $repoRoot 'dist'

function Write-ShellScript {
    param(
        [string]$Path,
        [string]$Content
    )

    $normalized = $Content -replace "`r?`n", "`n"
    [System.IO.File]::WriteAllText($Path, $normalized, [System.Text.Encoding]::UTF8)
}

foreach ($rid in $Runtime) {
    $runtimeRoot = Join-Path $distRoot $rid
    $targetDir = Join-Path $runtimeRoot 'ExcelMcp.Server'

    Write-Host "Cleaning output at $targetDir" -ForegroundColor Cyan
    if (Test-Path $targetDir) {
        Remove-Item $targetDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $targetDir | Out-Null

    $publishArgs = @(
        '-c', $Configuration,
        '-r', $rid,
        '--self-contained', 'true',
        '/p:PublishSingleFile=true',
        '/p:IncludeNativeLibrariesForSelfExtract=true',
        '/p:EnableCompressionInSingleFile=true',
        '/p:PublishTrimmed=false'
    )

    Write-Host "Publishing ExcelMcp.Server ($rid)" -ForegroundColor Cyan
    & dotnet publish (Join-Path $repoRoot 'src/ExcelMcp.Server/ExcelMcp.Server.csproj') @publishArgs -o $targetDir

    $readmePath = Join-Path $targetDir 'README.txt'
    $exeName = if ($rid -like 'win-*') { 'ExcelMcp.Server.exe' } else { 'ExcelMcp.Server' }
    $readme = @"
Excel MCP Server Bundle ($rid)
==============================

Contents
--------
- $exeName — self-contained single-file executable
- run-server.ps1 / run-server.sh — prompts for a workbook path and starts the server

Usage (Windows PowerShell)
--------------------------
1. Run .\run-server.ps1
2. Provide the Excel workbook path when prompted (e.g. C:\Data\sample.xlsx)
3. The server stays active until you press Ctrl+C

Usage (Linux / macOS bash)
--------------------------
1. chmod +x run-server.sh (once)
2. ./run-server.sh --workbook /path/to/workbook.xlsx
3. Stop with Ctrl+C when finished

Notes
-----
- The executable can also be launched directly; if no workbook is supplied it will prompt interactively.
- Ensure the workbook file is readable by the account running the server.
"@
    Set-Content -Path $readmePath -Value $readme -Encoding UTF8

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

$resolvedWorkbook = (Resolve-Path $WorkbookPath).Path
$serverExe = Join-Path $PSScriptRoot '__EXE_NAME__'

if (-not (Test-Path $serverExe)) {
    Write-Error "Server executable not found at $serverExe"
    exit 1
}

Write-Host "Starting Excel MCP server for $resolvedWorkbook" -ForegroundColor Cyan
& $serverExe --workbook $resolvedWorkbook
'@
    $serverLauncherPs1 = $serverLauncherPs1.Replace('__EXE_NAME__', $exeName)
    Set-Content -Path (Join-Path $targetDir 'run-server.ps1') -Value $serverLauncherPs1 -Encoding UTF8

    $serverLauncherBat = '@echo off`r`npwsh -ExecutionPolicy Bypass -File "%~dp0run-server.ps1" %*`r`n'
    Set-Content -Path (Join-Path $targetDir 'run-server.bat') -Value $serverLauncherBat -Encoding ASCII

    $serverLauncherSh = @'
#!/usr/bin/env bash
set -euo pipefail

SCRIPT_SOURCE="${BASH_SOURCE[0]}"
if [[ "${SCRIPT_SOURCE}" != /* ]]; then
    SCRIPT_SOURCE="${PWD}/${SCRIPT_SOURCE}"
fi
SCRIPT_DIR="${SCRIPT_SOURCE%/*}"
SERVER_EXE="${SCRIPT_DIR}/__EXE_NAME__"

if [[ ! -x "$SERVER_EXE" && -f "$SERVER_EXE" ]]; then
    chmod +x "$SERVER_EXE" || true
fi

if [[ ! -x "$SERVER_EXE" ]]; then
    echo "Server executable not found at $SERVER_EXE" >&2
    exit 1
fi

WORKBOOK=""

while [[ $# -gt 0 ]]; do
    case "$1" in
        -w|--workbook)
            WORKBOOK="$2"
            shift 2
            ;;
        --workbook=*)
            WORKBOOK="${1#*=}"
            shift
            ;;
        *)
            echo "Unknown argument: $1" >&2
            exit 1
            ;;
    esac
done

if [[ -z "$WORKBOOK" ]]; then
    read -r -p "Enter the full path to the Excel workbook: " WORKBOOK
fi

if [[ ! -f "$WORKBOOK" ]]; then
    echo "Workbook not found: $WORKBOOK" >&2
    exit 1
fi

exec "$SERVER_EXE" --workbook "$WORKBOOK"
'@
    $serverLauncherSh = $serverLauncherSh.Replace('__EXE_NAME__', $exeName)
    Write-ShellScript -Path (Join-Path $targetDir 'run-server.sh') -Content $serverLauncherSh

    if (-not $SkipZip) {
        $zipName = "excel-mcp-server-$rid.zip"
        $zipPath = Join-Path $distRoot $zipName
        Write-Host "Creating archive $zipPath" -ForegroundColor Cyan
        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force
        }
        Compress-Archive -Path (Join-Path $targetDir '*') -DestinationPath $zipPath
        Write-Host "Archive created: $zipPath" -ForegroundColor Green
    }

    Write-Host "Package complete: $targetDir" -ForegroundColor Green
}
