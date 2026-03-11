#!/usr/bin/env bash
set -euo pipefail

# Installs repo-local git hooks for EJS reminders.
# This sets the hooks path to .githooks for this repository.

repo_root="$(git rev-parse --show-toplevel 2>/dev/null || true)"
if [[ -z "$repo_root" ]]; then
  echo "EJS: not inside a git repository."
  exit 1
fi

cd "$repo_root"

git config core.hooksPath .githooks

chmod +x .githooks/post-commit .githooks/pre-push 2>/dev/null || true

echo "EJS: installed git hooks (core.hooksPath=.githooks)."
echo "EJS: hooks run on commit and push to remind you to capture a Session Journey."