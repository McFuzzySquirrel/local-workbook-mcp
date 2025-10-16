<#
.SYNOPSIS
    Builds a Windows Installer (MSI) for the Excel MCP tools using WiX.

.DESCRIPTION
    Restores the WiX tool (via dotnet tool), publishes self-contained binaries for the
    specified runtime using scripts/package.ps1, and then invokes wix build to produce
    an MSI that installs the published bundle into Program Files and adds Start Menu shortcuts.

.EXAMPLE
    pwsh -File scripts/package-wix.ps1

.EXAMPLE
    pwsh -File scripts/package-wix.ps1 -Runtime win-arm64
#>
[CmdletBinding()]
param(
    [string]$Runtime = "win-x64",
    [string]$Configuration = "Release"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$distRoot = Join-Path $repoRoot "dist"
$publishedDir = Join-Path $distRoot $Runtime
$msiName = "excel-mcp-$Runtime.msi"
$msiPath = Join-Path $distRoot $msiName
$wixProject = Join-Path $repoRoot "installer/wix/ExcelMcp.Installer.wixproj"

Write-Host "Building self-contained bundle for $Runtime ($Configuration)" -ForegroundColor Cyan
pwsh -File (Join-Path $repoRoot "scripts/package.ps1") -Runtime $Runtime -Configuration $Configuration

Write-Host "Restoring WiX tool" -ForegroundColor Cyan
& dotnet tool restore | Out-Null

if (-not (Test-Path $publishedDir)) {
    throw "Published directory not found: $publishedDir"
}

Write-Host "Building MSI $msiPath" -ForegroundColor Cyan
$buildArgs = @(
    "build",
    $wixProject,
    "-c",
    $Configuration,
    "-p:PublishedDir=$publishedDir",
    "-p:TargetName=$($msiName.Substring(0, $msiName.Length - 4))",
    "-p:OutputPath=$distRoot"
)
& dotnet @buildArgs

if (-not (Test-Path $msiPath)) {
    throw "MSI build succeeded but $msiPath was not created."
}

Write-Host "MSI generated: $msiPath" -ForegroundColor Green
