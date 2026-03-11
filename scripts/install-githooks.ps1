# Installs repo-local git hooks for EJS reminders (Windows-friendly).
# Sets core.hooksPath to .githooks for this repository.

$ErrorActionPreference = 'Stop'

try {
  $repoRoot = (git rev-parse --show-toplevel 2>$null).Trim()
} catch {
  Write-Error "EJS: not inside a git repository."
  exit 1
}

if ([string]::IsNullOrWhiteSpace($repoRoot)) {
  Write-Error "EJS: not inside a git repository."
  exit 1
}

Set-Location $repoRoot

git config core.hooksPath .githooks

Write-Host "EJS: installed git hooks (core.hooksPath=.githooks)."
Write-Host "EJS: hooks run on commit and push to remind you to capture a Session Journey."