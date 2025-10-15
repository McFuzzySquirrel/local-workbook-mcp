Param(
  [string]$Runtime = "win-x64",
  [string]$Configuration = "Release",
  [string]$OutDir = "dist/adapter"
)

$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath (Split-Path -Parent $MyInvocation.MyCommand.Path) | Out-Null
Set-Location .. | Out-Null

dotnet publish .\ExcelMcp.Adapter\ExcelMcp.Adapter.csproj `
  -c $Configuration `
  -r $Runtime `
  -o $OutDir `
  /p:PublishSingleFile=true `
  /p:PublishTrimmed=true `
  /p:IncludeNativeLibrariesForSelfExtract=true

Write-Host "Published to: $(Resolve-Path $OutDir)"