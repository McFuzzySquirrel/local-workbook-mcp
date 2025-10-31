<#
.SYNOPSIS
    Publishes the Excel MCP Semantic Kernel Agent as a self-contained, single-file executable bundle.

.DESCRIPTION
    For each requested Runtime Identifier (RID) the script performs a self-contained
    publish of ExcelMcp.SkAgent, stages the output under dist/<rid>/ExcelMcp.SkAgent,
    and adds launch helpers so end users can run the terminal agent interactively.
    Optionally, a zip archive is produced per RID for easy distribution.

.EXAMPLE
    pwsh -File scripts/package-skagent.ps1

.EXAMPLE
    pwsh -File scripts/package-skagent.ps1 -Runtime @('win-x64','linux-x64')

.EXAMPLE
    pwsh -File scripts/package-skagent.ps1 -Runtime win-x64 -SkipZip
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
    $targetDir = Join-Path $runtimeRoot 'ExcelMcp.SkAgent'

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

    Write-Host "Publishing ExcelMcp.SkAgent ($rid)" -ForegroundColor Cyan
    & dotnet publish (Join-Path $repoRoot 'src/ExcelMcp.SkAgent/ExcelMcp.SkAgent.csproj') @publishArgs -o $targetDir

    $readmePath = Join-Path $targetDir 'README.txt'
    $exeName = if ($rid -like 'win-*') { 'ExcelMcp.SkAgent.exe' } else { 'ExcelMcp.SkAgent' }
    $readme = @"
Excel MCP Semantic Kernel Agent Bundle ($rid)
==============================================

Contents
--------
- $exeName — self-contained single-file executable
- run-skagent.ps1 / run-skagent.sh / run-skagent.bat — prompts for workbook path and launches the agent

Quick Start (Windows PowerShell)
--------------------------------
1. Start your local LLM server (e.g., LM Studio on http://localhost:1234)
2. Run .\run-skagent.ps1
3. Provide the Excel workbook path when prompted
4. Chat with your workbook using natural language!

Quick Start (Linux / macOS bash)
--------------------------------
1. Start your local LLM server (e.g., Ollama or LM Studio)
2. chmod +x run-skagent.sh (once)
3. ./run-skagent.sh --workbook /path/to/workbook.xlsx
4. Type your questions and press Enter

Environment Variables
---------------------
LLM_BASE_URL     - LLM endpoint URL (default: http://localhost:1234)
LLM_MODEL_ID     - Model name (default: phi-4-mini-reasoning)
LLM_API_KEY      - API key (default: not-used)
EXCEL_MCP_WORKBOOK - Workbook path to skip prompts

Available Commands (in the agent REPL)
---------------------------------------
help, ?      - Show help
clear, cls   - Clear screen
exit, quit, q - Exit
<anything else> - Ask about the workbook

Example Queries
---------------
"What sheets are in this workbook?"
"Show me the first 10 rows of the Sales table"
"Search for 'laptop' across all sheets"
"How many rows are in the Products table?"

Notes
-----
- The agent uses Semantic Kernel to automatically call Excel tools
- All processing happens locally - no data leaves your machine
- Requires a running LLM server compatible with OpenAI API
"@
    Set-Content -Path $readmePath -Value $readme -Encoding UTF8

    $agentLauncherPs1 = @'
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath
)

if (-not $WorkbookPath) {
    $WorkbookPath = $env:EXCEL_MCP_WORKBOOK
}

if (-not $WorkbookPath) {
    $WorkbookPath = Read-Host 'Enter the full path to the Excel workbook'
}

if (-not $WorkbookPath) {
    Write-Error 'A workbook path is required.'
    exit 1
}

if (-not (Test-Path $WorkbookPath)) {
    Write-Error "Workbook not found: $WorkbookPath"
    exit 1
}

$resolvedWorkbook = (Resolve-Path $WorkbookPath).Path

$agentExe = Join-Path $PSScriptRoot '__EXE_NAME__'
if (-not (Test-Path $agentExe)) {
    Write-Error "Agent executable not found at $agentExe"
    exit 1
}

Write-Host "Starting Excel MCP Semantic Kernel Agent..." -ForegroundColor Cyan
Write-Host "Make sure your LLM server is running (e.g., LM Studio on http://localhost:1234)" -ForegroundColor Yellow
Write-Host ""

& $agentExe --workbook $resolvedWorkbook
'@
    $agentLauncherPs1 = $agentLauncherPs1.Replace('__EXE_NAME__', $exeName)
    Set-Content -Path (Join-Path $targetDir 'run-skagent.ps1') -Value $agentLauncherPs1 -Encoding UTF8

    $agentLauncherBat = '@echo off`r`npwsh -ExecutionPolicy Bypass -File "%~dp0run-skagent.ps1" %*`r`n'
    Set-Content -Path (Join-Path $targetDir 'run-skagent.bat') -Value $agentLauncherBat -Encoding ASCII

    $agentLauncherSh = @'
#!/usr/bin/env bash
set -euo pipefail

SCRIPT_SOURCE="${BASH_SOURCE[0]}"
if [[ "${SCRIPT_SOURCE}" != /* ]]; then
    SCRIPT_SOURCE="${PWD}/${SCRIPT_SOURCE}"
fi
SCRIPT_DIR="${SCRIPT_SOURCE%/*}"
AGENT_EXE="${SCRIPT_DIR}/__EXE_NAME__"

if [[ ! -x "${AGENT_EXE}" && -f "${AGENT_EXE}" ]]; then
    chmod +x "${AGENT_EXE}" || true
fi

if [[ ! -x "${AGENT_EXE}" ]]; then
    echo "Agent executable not found at ${AGENT_EXE}" >&2
    exit 1
fi

WORKBOOK="${EXCEL_MCP_WORKBOOK:-}"

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
            echo "Unknown option: $1" >&2
            exit 1
            ;;
    esac
done

if [[ -z "${WORKBOOK}" ]]; then
    read -r -p "Enter the full path to the Excel workbook: " WORKBOOK
fi

if [[ -z "${WORKBOOK}" ]]; then
    echo "A workbook path is required." >&2
    exit 1
fi

if [[ ! -f "${WORKBOOK}" ]]; then
    echo "Workbook not found: ${WORKBOOK}" >&2
    exit 1
fi

echo "Starting Excel MCP Semantic Kernel Agent..."
echo "Make sure your LLM server is running (e.g., LM Studio or Ollama)"
echo ""

exec "${AGENT_EXE}" --workbook "${WORKBOOK}"
'@
    $agentLauncherSh = $agentLauncherSh.Replace('__EXE_NAME__', $exeName)
    Write-ShellScript -Path (Join-Path $targetDir 'run-skagent.sh') -Content $agentLauncherSh

    if (-not $SkipZip) {
        $zipName = "excel-mcp-skagent-$rid.zip"
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
