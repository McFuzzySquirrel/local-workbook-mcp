# Web Chat Agent - Linux Compatibility Roadmap

**Date:** November 1, 2025  
**Platform:** Raspberry Pi (aarch64) / Debian 12 / .NET 9.0  
**Status:** ✅ Core components complete, needs testing & Linux path fixes

---

## Current Status

### ✅ What's Working
- All .NET projects build successfully on Linux (arm64)
- ExcelMcp.ChatWeb compiles without errors
- Unit tests pass (3/3)
- CLI agent (SkAgent) fully functional
- All Phase 1-2 tasks complete (Setup & Foundation)
- Most Phase 3 tasks complete (User Story 1)

### ⚠️ Issues to Fix

1. **Windows Path in appsettings.json**
   ```json
   "ExcelMcp": {
     "ServerPath": "D:\\GitHub Projects\\..."  // ❌ Windows path
   }
   ```
   
2. **Missing Linux-specific configuration**
   - Need `appsettings.Development.json` for Linux paths
   - MCP server path needs to be Linux-compatible

3. **Validation tasks not completed** (T078-T084)
   - Manual testing not done
   - Integration testing needed

---

## Action Plan

### PHASE 1: Fix Linux Configuration (15 minutes)

#### Task 1.1: Create Linux-friendly development config
Create `src/ExcelMcp.ChatWeb/appsettings.Development.json`:

```json
{
  "LlmStudio": {
    "BaseUrl": "http://localhost:1234",
    "Model": "qwen2.5-1.5b-instruct"
  },
  "ExcelMcp": {
    "ServerPath": "/home/mcsquirrel/Github Projects/local-workbook-mcp/src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server"
  },
  "SemanticKernel": {
    "Model": "qwen2.5-1.5b-instruct",
    "BaseUrl": "http://localhost:1234/v1",
    "ApiKey": "not-needed-for-local",
    "TimeoutSeconds": 480
  }
}
```

#### Task 1.2: Build MCP Server
```bash
dotnet build src/ExcelMcp.Server -c Debug
# Verify it exists
ls -lh src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server
```

#### Task 1.3: Verify executable permissions
```bash
chmod +x src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server
```

---

### PHASE 2: Test Basic Functionality (30 minutes)

#### Task 2.1: Start LLM Server
Ensure you have a local LLM running on port 1234:
- LM Studio with qwen2.5-1.5b-instruct model
- OR Ollama with compatible model
- OR any OpenAI-compatible endpoint

```bash
# Check if LLM is running
curl http://localhost:1234/v1/models
```

#### Task 2.2: Run ChatWeb Application
```bash
cd /home/mcsquirrel/Github\ Projects/local-workbook-mcp
ASPNETCORE_ENVIRONMENT=Development dotnet run --project src/ExcelMcp.ChatWeb
```

Expected output:
```
info: Microsoft.Hosting.Lifetime[14]
      Now listening on: http://localhost:5000
```

#### Task 2.3: Manual Testing Checklist

Open browser to `http://localhost:5000`

**T078 - Workbook Load Flow:**
- [ ] Click "Browse" and select `test-data/ProjectTracking.xlsx`
- [ ] Verify workbook metadata appears in sidebar (sheet count, sheet names)
- [ ] Verify success message: "✅ Workbook loaded successfully"

**T079 - Basic Query:**
- [ ] Ask: "What sheets are in this workbook?"
- [ ] Verify agent lists: "Tasks", "Projects", "TimeLog", "Users"
- [ ] Check debug log shows tool call to `get_workbook_summary`

**T080 - Data Retrieval:**
- [ ] Ask: "Show me the first 10 rows of the Tasks table"
- [ ] Verify HTML table is rendered with proper headers
- [ ] Check that TaskID, Title, Priority, Status columns appear

**T081 - Error Handling (No Workbook):**
- [ ] Clear history and reload page
- [ ] Ask a query WITHOUT loading workbook
- [ ] Verify error: "Please load a workbook first"

**T082 - Error Handling (Bad File):**
- [ ] Try to load a non-Excel file (e.g., README.md)
- [ ] Verify sanitized error message appears
- [ ] Check that correlationId is logged

**T083 - Query Timeout:**
- [ ] Ask a complex query (e.g., "Search for 'test' across all sheets")
- [ ] Verify query completes within 30 seconds OR shows timeout error
- [ ] Check retry option is presented

**T084 - End-to-End Validation:**
- [ ] Load workbook → ask 3 different questions → verify all work
- [ ] Switch to different workbook (BudgetTracker.xlsx)
- [ ] Verify context switches correctly
- [ ] Check conversation history shows workbook switch marker

---

### PHASE 3: Linux-Specific Enhancements (1 hour)

#### Task 3.1: Add Linux-specific package script
Create `scripts/package-chatweb-linux.ps1`:

```powershell
#!/usr/bin/env pwsh
# Package ChatWeb for Linux (arm64/x64)

param(
    [string]$Runtime = "linux-arm64",  # or linux-x64
    [switch]$SkipZip
)

$ProjectPath = "src/ExcelMcp.ChatWeb/ExcelMcp.ChatWeb.csproj"
$OutputDir = "dist/$Runtime/ExcelMcp.ChatWeb"

Write-Host "📦 Packaging ExcelMcp.ChatWeb for $Runtime..." -ForegroundColor Cyan

# Clean previous build
Remove-Item -Path $OutputDir -Recurse -Force -ErrorAction SilentlyContinue

# Publish single-file executable
dotnet publish $ProjectPath `
    --configuration Release `
    --runtime $Runtime `
    --self-contained true `
    --output $OutputDir `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true

if ($LASTEXITCODE -ne 0) {
    Write-Error "❌ Publish failed"
    exit 1
}

# Create run script
$RunScript = @"
#!/bin/bash
cd "\$(dirname "\$0")"
export ASPNETCORE_ENVIRONMENT=Production
./ExcelMcp.ChatWeb "\$@"
"@

Set-Content -Path "$OutputDir/run-chatweb.sh" -Value $RunScript
chmod +x "$OutputDir/run-chatweb.sh"
chmod +x "$OutputDir/ExcelMcp.ChatWeb"

Write-Host "✅ Package created: $OutputDir" -ForegroundColor Green
```

#### Task 3.2: Test packaged build
```bash
pwsh -File scripts/package-chatweb-linux.ps1 -Runtime linux-arm64

# Test the packaged version
cd dist/linux-arm64/ExcelMcp.ChatWeb
./run-chatweb.sh
```

#### Task 3.3: Document Linux setup
Update `docs/UserGuide.md` with Linux-specific instructions:
- Required dependencies (libgdiplus for Excel processing)
- Setting up LLM on Raspberry Pi
- Performance expectations on ARM

---

### PHASE 4: Performance Tuning for Raspberry Pi (Optional)

Since you're on Raspberry Pi (limited resources):

#### Task 4.1: Optimize memory usage
- Set smaller model (phi-2 or qwen2.5-0.5b)
- Reduce MaxContextTurns to 10
- Limit MaxResponseLength to 5000

#### Task 4.2: Add resource monitoring
```bash
# Monitor during testing
top -p $(pgrep -f ExcelMcp.ChatWeb)
```

#### Task 4.3: Benchmark
- Test with 50MB workbook (SC-001: should load in <5 seconds)
- Measure query response time (SC-002: <10 seconds for structure queries)

---

## Linux Compatibility Checklist

### System Requirements
- [ ] .NET 9.0 SDK installed ✅ (confirmed)
- [ ] libgdiplus for Excel processing (`sudo apt install libgdiplus`)
- [ ] Local LLM server (LM Studio/Ollama)
- [ ] Minimum 2GB RAM available for ChatWeb

### Code Changes Needed
- [ ] Fix appsettings.json server path (Task 1.1)
- [ ] Create Linux packaging script (Task 3.1)
- [ ] Update documentation (Task 3.3)

### Testing on Linux
- [ ] ChatWeb builds successfully ✅
- [ ] ChatWeb runs without errors
- [ ] Can load Excel workbooks
- [ ] Can query workbooks via UI
- [ ] MCP server starts correctly
- [ ] Blazor UI renders properly
- [ ] File selector works on Linux

---

## Quick Start (Linux)

### Prerequisites
```bash
# Install required libraries
sudo apt update
sudo apt install libgdiplus

# Verify .NET 9
dotnet --version  # Should be 9.0.x
```

### Run Development Server
```bash
cd /home/mcsquirrel/Github\ Projects/local-workbook-mcp

# Start your LLM (e.g., LM Studio on port 1234)

# Run ChatWeb
ASPNETCORE_ENVIRONMENT=Development \
EXCEL_MCP_SERVER="$(pwd)/src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server" \
dotnet run --project src/ExcelMcp.ChatWeb

# Open browser to http://localhost:5000
```

---

## Known Issues & Workarounds

### Issue: File picker doesn't work in browser
**Workaround:** Use pre-created sample workbooks in `test-data/` folder

### Issue: Slow performance on Raspberry Pi
**Workaround:** 
- Use smaller models (qwen2.5-0.5b or phi-2)
- Reduce context window to 10 turns
- Test with smaller workbooks (<10MB)

### Issue: MCP server not found
**Workaround:**
```bash
# Build server first
dotnet build src/ExcelMcp.Server

# Set environment variable
export EXCEL_MCP_SERVER="$(pwd)/src/ExcelMcp.Server/bin/Debug/net9.0/ExcelMcp.Server"
```

---

## Next Steps Priority

1. **IMMEDIATE:** Fix appsettings.Development.json (5 min)
2. **TODAY:** Complete manual testing (T078-T084) (30 min)
3. **THIS WEEK:** Create Linux packaging script (1 hr)
4. **NICE TO HAVE:** Performance tuning for Pi (2 hrs)

---

## Success Criteria (from spec.md)

Track these on Raspberry Pi:
- [ ] SC-001: Workbook load < 5 seconds for 50MB files
- [ ] SC-002: Basic queries < 10 seconds response time
- [ ] SC-003: 90% accuracy on structure queries
- [ ] SC-006: Handle 100 sheets, 100k rows without crashing
- [ ] SC-010: Runs on 8GB RAM, standard CPU ✅ (Pi has this)

**Note:** Some performance metrics may be adjusted for ARM hardware.
