<#
.SYNOPSIS
    Publishes the Excel MCP chat web app as a self-contained bundle with static assets.

.DESCRIPTION
    For each requested Runtime Identifier (RID) the script performs a self-contained
    publish of ExcelMcp.ChatWeb, stages the output under dist/<rid>/ExcelMcp.ChatWeb,
    and adds launch helpers so end users can provide workbook/server paths interactively.
    The publish output always includes the wwwroot static assets required by the app.
    Optionally, a zip archive is produced per RID for easy distribution.

.EXAMPLE
    pwsh -File scripts/package-chatweb.ps1

.EXAMPLE
    pwsh -File scripts/package-chatweb.ps1 -Runtime @('win-x64','linux-x64')

.EXAMPLE
    pwsh -File scripts/package-chatweb.ps1 -Runtime win-x64 -SkipZip
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
$projectPath = Join-Path $repoRoot 'src/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb.csproj'
$projectRoot = Split-Path $projectPath

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
    $targetDir = Join-Path $runtimeRoot 'ExcelMcp.ChatWeb'

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

    Write-Host "Publishing ExcelMcp.ChatWeb ($rid)" -ForegroundColor Cyan
    & dotnet publish $projectPath @publishArgs -o $targetDir

    $wwwrootTarget = Join-Path $targetDir 'wwwroot'
    if (-not (Test-Path $wwwrootTarget)) {
        Write-Host "wwwroot not found in publish output; copying from project." -ForegroundColor Yellow
        Copy-Item -Path (Join-Path $projectRoot 'wwwroot') -Destination $wwwrootTarget -Recurse -Force
    }

    $readmePath = Join-Path $targetDir 'README.txt'
    $exeName = if ($rid -like 'win-*') { 'ExcelMcp.ChatWeb.exe' } else { 'ExcelMcp.ChatWeb' }
    $readme = @"
Excel MCP Chat Web Bundle ($rid)
================================

Contents
--------
- $exeName — self-contained single-file ASP.NET application
- wwwroot/ — static web assets served by the app
- run-chatweb.ps1 / run-chatweb.sh — prompts for workbook/server paths and optional URL binding

Quick Start (Windows PowerShell)
--------------------------------
1. Run .\run-chatweb.ps1
2. Provide the Excel workbook path when prompted (e.g. C:\Data\sample.xlsx)
3. Provide the Excel MCP server executable path if one is not detected automatically
4. Accept the default hosting URL (http://localhost:5000) or provide a custom binding

Quick Start (Linux / macOS bash)
--------------------------------
1. chmod +x run-chatweb.sh (once)
2. ./run-chatweb.sh --workbook /path/to/workbook.xlsx --server /path/to/ExcelMcp.Server --urls http://localhost:5000
3. Open the reported URL in a browser to access the chat interface

Notes
-----
- The app will launch the Excel MCP server process specified by --server.
- Set EXCEL_MCP_WORKBOOK or EXCEL_MCP_SERVER environment variables to skip prompts.
- Use --urls to control the binding (default http://localhost:5000).
"@
    Set-Content -Path $readmePath -Value $readme -Encoding UTF8

    $launcherPs1 = @'
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath,
    [string]$ServerPath,
    [string]$Urls,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Remaining
)

function Get-ServerCandidate {
    param([string]$BaseDir)

    $relativePaths = @(
        'ExcelMcp.Server.exe',
        'ExcelMcp.Server',
        'ExcelMcp.Server\ExcelMcp.Server.exe',
        'ExcelMcp.Server\ExcelMcp.Server',
        '..\ExcelMcp.Server\ExcelMcp.Server.exe',
        '..\ExcelMcp.Server\ExcelMcp.Server',
        '..\..\ExcelMcp.Server\ExcelMcp.Server.exe',
        '..\..\ExcelMcp.Server\ExcelMcp.Server'
    )

    foreach ($relative in $relativePaths) {
        try {
            $candidate = Join-Path $BaseDir $relative
            $full = [System.IO.Path]::GetFullPath($candidate)
        } catch {
            continue
        }

        if (Test-Path $full) {
            return (Resolve-Path $full).Path
        }
    }

    return $null
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

if (-not $ServerPath) {
    $defaultServer = Get-ServerCandidate -BaseDir $PSScriptRoot
    if ($defaultServer) {
        $input = Read-Host "Enter the Excel MCP server executable path [$defaultServer]"
        if ([string]::IsNullOrWhiteSpace($input)) {
            $ServerPath = $defaultServer
        } else {
            if (-not (Test-Path $input)) {
                Write-Error "Server executable not found: $input"
                exit 1
            }
            $ServerPath = (Resolve-Path $input).Path
        }
    } else {
        $ServerPath = Read-Host 'Enter the full path to the Excel MCP server executable'
        if (-not $ServerPath) {
            Write-Error 'A server executable path is required.'
            exit 1
        }
        if (-not (Test-Path $ServerPath)) {
            Write-Error "Server executable not found: $ServerPath"
            exit 1
        }
        $ServerPath = (Resolve-Path $ServerPath).Path
    }
}

if (-not $ServerPath) {
    Write-Error 'A server executable path is required.'
    exit 1
}

if (-not (Test-Path $ServerPath)) {
    Write-Error "Server executable not found: $ServerPath"
    exit 1
}

$resolvedServer = (Resolve-Path $ServerPath).Path

if (-not $Urls) {
    $Urls = Read-Host 'Enter the binding URL for the web app [http://localhost:5000]'
    if ([string]::IsNullOrWhiteSpace($Urls)) {
        $Urls = 'http://localhost:5000'
    }
}

$env:EXCEL_MCP_WORKBOOK = $resolvedWorkbook
$env:EXCEL_MCP_SERVER = $resolvedServer
$env:ASPNETCORE_URLS = $Urls

$chatExe = Join-Path $PSScriptRoot '__EXE_NAME__'
if (-not (Test-Path $chatExe)) {
    Write-Error "Chat web executable not found at $chatExe"
    exit 1
}

$arguments = @($Remaining)
if (-not $arguments) {
    $arguments = @('--urls', $Urls)
}

Write-Host "Starting Excel MCP chat web app on $Urls" -ForegroundColor Cyan
& $chatExe @arguments
'@
    $launcherPs1 = $launcherPs1.Replace('__EXE_NAME__', $exeName)
    Set-Content -Path (Join-Path $targetDir 'run-chatweb.ps1') -Value $launcherPs1 -Encoding UTF8

    $launcherBat = '@echo off`r`npwsh -ExecutionPolicy Bypass -File "%~dp0run-chatweb.ps1" %*`r`n'
    Set-Content -Path (Join-Path $targetDir 'run-chatweb.bat') -Value $launcherBat -Encoding ASCII

    $launcherSh = @'
#!/usr/bin/env bash
set -euo pipefail

SCRIPT_SOURCE="${BASH_SOURCE[0]}"
if [[ "${SCRIPT_SOURCE}" != /* ]]; then
    SCRIPT_SOURCE="${PWD}/${SCRIPT_SOURCE}"
fi
SCRIPT_DIR="${SCRIPT_SOURCE%/*}"
CHAT_EXE="${SCRIPT_DIR}/__EXE_NAME__"

if [[ ! -x "${CHAT_EXE}" && -f "${CHAT_EXE}" ]]; then
    chmod +x "${CHAT_EXE}" || true
fi

if [[ ! -x "${CHAT_EXE}" ]]; then
    echo "Chat web executable not found at ${CHAT_EXE}" >&2
    exit 1
fi

WORKBOOK=""
SERVER=""
URLS=""
REMAINING=()

find_server_executable() {
    local base="${SCRIPT_DIR}"
    local parent="${base%/*}"
    local grand="${parent%/*}"

    local candidates=(
        "${base}/ExcelMcp.Server"
        "${base}/ExcelMcp.Server.exe"
        "${base}/ExcelMcp.Server/ExcelMcp.Server"
        "${base}/ExcelMcp.Server/ExcelMcp.Server.exe"
        "${parent}/ExcelMcp.Server/ExcelMcp.Server"
        "${parent}/ExcelMcp.Server/ExcelMcp.Server.exe"
        "${grand}/ExcelMcp.Server/ExcelMcp.Server"
        "${grand}/ExcelMcp.Server/ExcelMcp.Server.exe"
    )

    for candidate in "${candidates[@]}"; do
        if [[ -f "${candidate}" ]]; then
            echo "${candidate}"
            return 0
        fi
    done

    return 1
}

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
        -s|--server)
            SERVER="$2"
            shift 2
            ;;
        --server=*)
            SERVER="${1#*=}"
            shift
            ;;
        -u|--urls)
            URLS="$2"
            shift 2
            ;;
        --urls=*)
            URLS="${1#*=}"
            shift
            ;;
        --)
            shift
            REMAINING+=("$@")
            break
            ;;
        -* )
            echo "Unknown option: $1" >&2
            exit 1
            ;;
        * )
            REMAINING+=("$1")
            shift
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

if [[ -z "${SERVER}" ]]; then
    DEFAULT_SERVER="$(find_server_executable || true)"
    if [[ -n "${DEFAULT_SERVER}" ]]; then
        read -r -p "Enter the Excel MCP server executable path [${DEFAULT_SERVER}]: " INPUT
        if [[ -z "${INPUT}" ]]; then
            SERVER="${DEFAULT_SERVER}"
        else
            SERVER="${INPUT}"
        fi
    else
        read -r -p "Enter the Excel MCP server executable path: " SERVER
    fi
fi

if [[ -z "${SERVER}" ]]; then
    echo "A server executable path is required." >&2
    exit 1
fi

if [[ ! -f "${SERVER}" ]]; then
    echo "Server executable not found: ${SERVER}" >&2
    exit 1
fi

if [[ -z "${URLS}" ]]; then
    read -r -p "Enter the binding URL for the web app [http://localhost:5000]: " URL_INPUT
    if [[ -z "${URL_INPUT}" ]]; then
        URLS="http://localhost:5000"
    else
        URLS="${URL_INPUT}"
    fi
fi

export EXCEL_MCP_WORKBOOK="${WORKBOOK}"
export EXCEL_MCP_SERVER="${SERVER}"
export ASPNETCORE_URLS="${URLS}"

if (( ${#REMAINING[@]} == 0 )); then
    REMAINING=("--urls" "${URLS}")
fi

exec "${CHAT_EXE}" "${REMAINING[@]}"
'@
    $launcherSh = $launcherSh.Replace('__EXE_NAME__', $exeName)
    Write-ShellScript -Path (Join-Path $targetDir 'run-chatweb.sh') -Content $launcherSh

    if (-not $SkipZip) {
        $zipName = "excel-mcp-chatweb-$rid.zip"
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
