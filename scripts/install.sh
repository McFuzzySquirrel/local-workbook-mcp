#!/usr/bin/env bash
# install.sh — Build and install ExcelMcp.Server and ExcelMcp.ChatWeb on Linux/macOS
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PUBLISH_DIR="${REPO_ROOT}/publish"
SERVER_PUBLISH="${PUBLISH_DIR}/server"
CHATWEB_PUBLISH="${PUBLISH_DIR}/chatweb"

echo "=== local-workbook-mcp install ==="
echo "Repo root : ${REPO_ROOT}"
echo "Output    : ${PUBLISH_DIR}"
echo ""

# Publish MCP server
echo "[1/2] Publishing ExcelMcp.Server..."
dotnet publish "${REPO_ROOT}/src/ExcelMcp.Server/ExcelMcp.Server.csproj" \
  --configuration Release \
  --output "${SERVER_PUBLISH}" \
  --nologo
chmod +x "${SERVER_PUBLISH}/ExcelMcp.Server"

# Publish Blazor ChatWeb (optional)
echo "[2/2] Publishing ExcelMcp.ChatWeb..."
dotnet publish "${REPO_ROOT}/src/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb.csproj" \
  --configuration Release \
  --output "${CHATWEB_PUBLISH}" \
  --nologo
chmod +x "${CHATWEB_PUBLISH}/ExcelMcp.ChatWeb"

echo ""
echo "Done! Binaries written to:"
echo "  Server  : ${SERVER_PUBLISH}/ExcelMcp.Server"
echo "  ChatWeb : ${CHATWEB_PUBLISH}/ExcelMcp.ChatWeb"
echo ""
echo "Quick-start:"
echo "  # Run MCP server standalone (for Claude Desktop / Cursor / VS Code)"
echo "  ${SERVER_PUBLISH}/ExcelMcp.Server --workbook /path/to/workbook.xlsx"
echo ""
echo "  # Run Blazor Chat UI (set EXCEL_MCP_SERVER env var to the server binary)"
echo "  EXCEL_MCP_SERVER=${SERVER_PUBLISH}/ExcelMcp.Server ${CHATWEB_PUBLISH}/ExcelMcp.ChatWeb"
echo ""
echo "See mcp-config/ for client config templates."
