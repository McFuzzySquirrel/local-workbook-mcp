<#
.SYNOPSIS
    Publishes the Excel MCP client as a self-contained, single-file executable bundle.

.DESCRIPTION
    For each requested Runtime Identifier (RID) the script performs a self-contained
    publish of ExcelMcp.Client, stages the output under dist/<rid>/ExcelMcp.Client,
    and adds launch helpers so end users can connect to a workbook and MCP server
    interactively. Optionally, a zip archive is produced per RID for easy distribution.

.EXAMPLE
    pwsh -File scripts/package-client.ps1

.EXAMPLE
    pwsh -File scripts/package-client.ps1 -Runtime @('win-x64','linux-x64')

.EXAMPLE
    pwsh -File scripts/package-client.ps1 -Runtime win-x64 -SkipZip
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
    $targetDir = Join-Path $runtimeRoot 'ExcelMcp.Client'

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

    Write-Host "Publishing ExcelMcp.Client ($rid)" -ForegroundColor Cyan
    & dotnet publish (Join-Path $repoRoot 'src/ExcelMcp.Client/ExcelMcp.Client.csproj') @publishArgs -o $targetDir

    $readmePath = Join-Path $targetDir 'README.txt'
    $exeName = if ($rid -like 'win-*') { 'ExcelMcp.Client.exe' } else { 'ExcelMcp.Client' }
    $readme = @"
Excel MCP Client Bundle ($rid)
===============================

Contents
--------
- $exeName — self-contained single-file executable
- run-client.ps1 / run-client.sh — prompts for workbook and server paths, then runs a command

Quick Start (Windows PowerShell)
--------------------------------
1. Run .\run-client.ps1
2. Provide the Excel workbook path when prompted (e.g. C:\Data\sample.xlsx)
3. Provide the Excel MCP server executable path if one is not detected automatically
4. Choose the client command to run (default is 'list')

Quick Start (Linux / macOS bash)
--------------------------------
1. chmod +x run-client.sh (once)
2. ./run-client.sh --workbook /path/to/workbook.xlsx --server /path/to/ExcelMcp.Server list
3. Add additional arguments after the command (e.g. --query "budget")

Notes
-----
- The client will start the specified MCP server executable as a child process.
- Set EXCEL_MCP_WORKBOOK or EXCEL_MCP_SERVER environment variables to skip prompts.
- Commands available: list, search, preview, resources.
"@
    Set-Content -Path $readmePath -Value $readme -Encoding UTF8

    $clientLauncherPs1 = @'
param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath,
    [string]$ServerPath,
    [string]$Command,
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$CommandArgs
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

if (-not $Command) {
    $Command = Read-Host 'Enter the client command to run (list/search/preview/resources) [list]'
    if ([string]::IsNullOrWhiteSpace($Command)) {
        $Command = 'list'
    }
}

if (-not $Command) {
    Write-Error 'A command is required.'
    exit 1
}

$env:EXCEL_MCP_WORKBOOK = $resolvedWorkbook
$env:EXCEL_MCP_SERVER = $resolvedServer

$clientExe = Join-Path $PSScriptRoot '__EXE_NAME__'
if (-not (Test-Path $clientExe)) {
    Write-Error "Client executable not found at $clientExe"
    exit 1
}

$arguments = @('--workbook', $resolvedWorkbook, '--server', $resolvedServer, $Command)
if ($CommandArgs) {
    $arguments += $CommandArgs
}

Write-Host "Starting Excel MCP client: $Command" -ForegroundColor Cyan
& $clientExe @arguments
'@
    $clientLauncherPs1 = $clientLauncherPs1.Replace('__EXE_NAME__', $exeName)
    Set-Content -Path (Join-Path $targetDir 'run-client.ps1') -Value $clientLauncherPs1 -Encoding UTF8

    $clientLauncherBat = '@echo off`r`npwsh -ExecutionPolicy Bypass -File "%~dp0run-client.ps1" %*`r`n'
    Set-Content -Path (Join-Path $targetDir 'run-client.bat') -Value $clientLauncherBat -Encoding ASCII

    $clientLauncherSh = @'
#!/usr/bin/env bash
set -euo pipefail

SCRIPT_SOURCE="${BASH_SOURCE[0]}"
if [[ "${SCRIPT_SOURCE}" != /* ]]; then
    SCRIPT_SOURCE="${PWD}/${SCRIPT_SOURCE}"
fi
SCRIPT_DIR="${SCRIPT_SOURCE%/*}"
CLIENT_EXE="${SCRIPT_DIR}/__EXE_NAME__"

if [[ ! -x "${CLIENT_EXE}" && -f "${CLIENT_EXE}" ]]; then
    chmod +x "${CLIENT_EXE}" || true
fi

if [[ ! -x "${CLIENT_EXE}" ]]; then
    echo "Client executable not found at ${CLIENT_EXE}" >&2
    exit 1
}

WORKBOOK=""
SERVER=""
COMMAND=""
COMMAND_ARGS=()

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
        -c|--command)
            COMMAND="$2"
            shift 2
            ;;
        --command=*)
            COMMAND="${1#*=}"
            shift
            ;;
        --)
            shift
            COMMAND_ARGS+=("$@")
            break
            ;;
        -* )
            echo "Unknown option: $1" >&2
            exit 1
            ;;
        *)
            COMMAND="$1"
            shift
            COMMAND_ARGS+=("$@")
            break
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

if [[ -z "${COMMAND}" ]]; then
    read -r -p "Enter the client command to run (list/search/preview/resources) [list]: " COMMAND
fi

if [[ -z "${COMMAND}" ]]; then
    COMMAND="list"
fi

export EXCEL_MCP_WORKBOOK="${WORKBOOK}"
export EXCEL_MCP_SERVER="${SERVER}"

set -- --workbook "${WORKBOOK}" --server "${SERVER}" "${COMMAND}"
if (( ${#COMMAND_ARGS[@]} > 0 )); then
    set -- "$@" "${COMMAND_ARGS[@]}"
fi

exec "${CLIENT_EXE}" "$@"
'@
    $clientLauncherSh = $clientLauncherSh.Replace('__EXE_NAME__', $exeName)
    Write-ShellScript -Path (Join-Path $targetDir 'run-client.sh') -Content $clientLauncherSh

    if (-not $SkipZip) {
        $zipName = "excel-mcp-client-$rid.zip"
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
