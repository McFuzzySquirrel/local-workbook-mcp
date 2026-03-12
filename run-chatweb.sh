#!/bin/bash
# Quick start script for ExcelMcp.ChatWeb on Linux

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "🚀 Starting ExcelMcp.ChatWeb on Linux..."
echo ""

# Check prerequisites
echo "✓ Checking prerequisites..."

# Check .NET
if ! command -v dotnet &> /dev/null; then
    echo "❌ .NET SDK not found. Please install .NET 10.0+"
    exit 1
fi
echo "  .NET version: $(dotnet --version)"

# Check libgdiplus (for Excel processing)
if ! ldconfig -p | grep -q libgdiplus; then
    echo "⚠️  libgdiplus not found. Excel processing may fail."
    echo "   Install with: sudo apt install libgdiplus"
fi

# Check if LLM is running
if ! curl -s http://localhost:11434/api/tags > /dev/null 2>&1; then
    echo "⚠️  LLM server not detected on http://localhost:11434"
    echo "   Make sure Ollama is running: ollama serve"
    echo "   Continuing anyway..."
fi

# Build MCP Server if needed
if [ ! -f "src/ExcelMcp.Server/bin/Debug/net10.0/ExcelMcp.Server" ]; then
    echo ""
    echo "📦 Building MCP Server..."
    dotnet build src/ExcelMcp.Server -c Debug --nologo
    chmod +x src/ExcelMcp.Server/bin/Debug/net10.0/ExcelMcp.Server
fi

# Set environment
export ASPNETCORE_ENVIRONMENT=Development
export EXCEL_MCP_SERVER="$(pwd)/src/ExcelMcp.Server/bin/Debug/net10.0/ExcelMcp.Server"

echo ""
echo "✅ Ready to start!"
echo ""
echo "📝 Configuration:"
echo "   Environment: $ASPNETCORE_ENVIRONMENT"
echo "   MCP Server: $EXCEL_MCP_SERVER"
echo "   LLM Endpoint: http://localhost:11434"
echo ""
echo "🌐 Starting web server..."
echo "   Open browser to: http://localhost:5000"
echo ""

# Run the application
dotnet run --project src/ExcelMcp.ChatWeb
