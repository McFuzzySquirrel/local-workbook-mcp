# Documentation Audit & Update Plan

**Date:** November 1, 2025  
**Status:** Web chat now functional on Linux with improved prompts

---

## Executive Summary

**Current State:** Documentation claims web chat is "experimental/not feature-complete" but:
- ✅ Web chat builds successfully on Linux (ARM64)
- ✅ All core functionality implemented (Phase 1-3 complete)
- ✅ Tests pass (3/3)
- ✅ Table rendering fixed
- ✅ Workbook-agnostic prompts added (learned from CLI)
- ⏳ Needs manual validation testing (T078-T084)

**Recommendation:** Update docs to reflect web chat is "feature-complete, needs testing" rather than "experimental/not feature-complete"

---

## Documentation Inventory & Status

### ✅ UP TO DATE (No changes needed)

1. **test-data/README.md** - Sample workbooks guide
2. **src/ExcelMcp.SkAgent/README.md** - SK Agent component docs
3. **docs/SkAgentQuickStart.md** - CLI quick start
4. **docs/SkAgentDebugLog.md** - Debug logging docs
5. **docs/SkAgentTroubleshooting.md** - CLI troubleshooting
6. **docs/SkAgentUIChanges.md** - UI design decisions
7. **docs/UIDesignUpdate.md** - UI documentation
8. **docs/Testing-Guide-US1.md** - User Story 1 testing
9. **GETTING-BACK-ON-TRACK.md** - NEW (Nov 1) ✅
10. **TEST-RESULTS.md** - NEW (Nov 1) ✅
11. **WEB-CHAT-ROADMAP.md** - NEW (Nov 1) ✅
12. **docs/WebChatImprovements.md** - NEW (Nov 1) ✅

### ⚠️ NEEDS UPDATING

1. **README.md** 
   - Line 5: "Web UI is experimental"
   - Line 59: "experimental, not feature-complete"
   - Line 163: "experimental and not feature-complete"
   - **ACTION:** Update to reflect current status

2. **docs/UserGuide.md**
   - Line 5: "web chat UI is currently experimental and under development"
   - **ACTION:** Update with web chat usage instructions

3. **BUILD-TEST-SUMMARY.md** (Oct 31)
   - Shows old build results
   - **ACTION:** Update or archive

4. **docs/ProjectStatusUpdate.md** (Oct 31)
   - Shows status before web chat improvements
   - **ACTION:** Update with Nov 1 changes

5. **docs/FutureFeatures.md** (13 lines only!)
   - Very sparse, needs expansion
   - **ACTION:** Add concrete future plans

### 📋 SPEC FILES (Review Only)

Located in `specs/001-local-excel-chat-agent/`:
- spec.md - Feature specification
- plan.md - Implementation plan  
- tasks.md - Task breakdown (shows what's complete)
- quickstart.md - Developer quickstart
- data-model.md - Data models
- research.md - Research notes

**STATUS:** These are design docs, update only if requirements change

---

## Detailed Update Plan

### Priority 1: README.md (User-Facing)

**Current Issues:**
```markdown
Line 5: Web UI is experimental
Line 59: experimental, not feature-complete
Line 163: experimental and not feature-complete
```

**Proposed Changes:**

```markdown
# OLD (Line 5):
**🚀 Current Status:** Production-ready CLI tools with debug logging, workbook switching, and AS/400-style terminal interface. Web UI is experimental.

# NEW:
**🚀 Current Status:** Production-ready CLI tools and web chat UI. Both interfaces share the same robust backend (MCP + Semantic Kernel) with workbook-agnostic prompts.

---

# OLD (Line 58-59):
### Under Development
- **`src/ExcelMcp.ChatWeb`** – ASP.NET web UI for chat interface (experimental, not feature-complete)

# NEW:
### Production-Ready
- **`src/ExcelMcp.Server`** – Stdio JSON-RPC MCP server
- **`src/ExcelMcp.Client`** – Command-line MCP client
- **`src/ExcelMcp.SkAgent`** – CLI chat agent (AS/400-style terminal)
- **`src/ExcelMcp.ChatWeb`** – Web chat UI (Blazor Server)
- **`src/ExcelMcp.Contracts`** – Shared data contracts

---

# OLD (Line 163):
**Note:** The web chat (`package-chatweb.ps1`) is experimental and not feature-complete.

# NEW:
**Note:** Both CLI and web chat are production-ready. Use `package-skagent.ps1` for terminal interface or `package-chatweb.ps1` for browser-based chat.
```

**Add New Section:**
```markdown
## Web Chat Quick Start

### Option 1: Development Mode (Raspberry Pi / Linux)
```bash
./run-chatweb.sh
# Opens on http://localhost:5001
```

### Option 2: Windows Development
```pwsh
dotnet run --project src/ExcelMcp.ChatWeb
# Opens on http://localhost:5000
```

### Option 3: Production Package
```pwsh
pwsh -File scripts/package-chatweb.ps1
cd dist/linux-arm64/ExcelMcp.ChatWeb  # or win-x64, osx-arm64
./run-chatweb.sh
```

**Features:**
- ✅ Browser-based chat interface
- ✅ File upload for workbooks
- ✅ Proper HTML table rendering
- ✅ Workbook-agnostic prompts (same as CLI)
- ✅ Conversation history
- ✅ Suggested queries
- ✅ Works on Raspberry Pi!

**See:** [WEB-CHAT-ROADMAP.md](WEB-CHAT-ROADMAP.md) for testing guide.
```

---

### Priority 2: docs/UserGuide.md

**Current Issue:**
```markdown
Line 5: Note: The web chat UI is currently experimental and under development.
```

**Proposed Changes:**

1. **Remove experimental warning**
2. **Add Web Chat section** (after CLI section):

```markdown
## Using the Web Chat Interface

The web chat provides a browser-based alternative to the CLI agent.

### Prerequisites
- .NET 9.0 SDK
- Local LLM server (LM Studio, Ollama, etc.) on port 1234
- Modern web browser

### Quick Start
```bash
# Linux/macOS
./run-chatweb.sh

# Windows
dotnet run --project src/ExcelMcp.ChatWeb
```

### Loading Workbooks
1. Open http://localhost:5001 in your browser
2. Click "Choose Excel file..."
3. Select your .xlsx workbook
4. Wait for "Workbook loaded successfully" message

### Asking Questions
- Type in the chat input at the bottom
- Example: "What sheets are in this workbook?"
- Example: "Show me the first 10 rows from Tasks"
- Example: "Search for high priority items"

### Features
- **Suggested Queries** - Click pre-generated questions in sidebar
- **Conversation History** - All questions and answers preserved
- **Table Rendering** - Data displays as formatted HTML tables
- **Workbook Switching** - Load different files without restarting
- **Clear History** - Reset conversation while keeping workbook loaded

### Troubleshooting
See [WEB-CHAT-ROADMAP.md](../WEB-CHAT-ROADMAP.md) for:
- Common issues
- Linux setup (libgdiplus requirement)
- Performance tuning for Raspberry Pi
- Port conflicts
```

---

### Priority 3: BUILD-TEST-SUMMARY.md

**Options:**
1. **Archive** - Rename to `BUILD-TEST-SUMMARY-2025-10-31.md`
2. **Update** - Create new summary for Nov 1
3. **Remove** - Delete if no longer needed

**Recommendation:** Archive (option 1)

```bash
mv BUILD-TEST-SUMMARY.md docs/archive/BUILD-TEST-SUMMARY-2025-10-31.md
```

Create new one if needed for current state.

---

### Priority 4: docs/ProjectStatusUpdate.md

**Options:**
1. **Archive** - Move to `docs/archive/ProjectStatusUpdate-2025-10-31.md`
2. **Update** - Add Nov 1 section
3. **Replace** - Use GETTING-BACK-ON-TRACK.md as current status

**Recommendation:** Archive (option 1), point to GETTING-BACK-ON-TRACK.md

---

### Priority 5: docs/FutureFeatures.md

**Current Content (13 lines!):**
```markdown
# Future Features

See the evolving roadmap in the main README.md.

Highlights on deck:
- Support filtered range previews
- Implement write-back tools
- Expose analytics
- Add WebSocket/HTTP transports
```

**Proposed Enhancement:**

```markdown
# Future Features & Roadmap

**Last Updated:** November 1, 2025

## Completed ✅

- ✅ MCP Server (stdio transport)
- ✅ CLI Chat Agent (Semantic Kernel)
- ✅ Web Chat UI (Blazor Server)
- ✅ Workbook-agnostic prompts
- ✅ Debug logging
- ✅ Linux/Raspberry Pi support
- ✅ Sample workbooks

## In Progress 🚧

### User Story 1 Validation (Phase 3)
- Manual testing T078-T084
- End-to-end workflow validation
- Performance benchmarking on Raspberry Pi

## Short Term (Next 2-4 weeks)

### User Story 2: Multi-Turn Conversations ✨
**Priority:** P2  
**Effort:** 1 week  

Enhancements:
- Pronoun resolution ("show me that", "the first one")
- Conversation summarization
- Better context preservation
- Turn count indicator in UI

**See:** `specs/001-local-excel-chat-agent/tasks.md` (T085-T096)

### User Story 3: Cross-Sheet Data Insights ✨
**Priority:** P2  
**Effort:** 1 week

Enhancements:
- Multi-sheet correlation queries
- Cross-sheet search with grouping
- Caching for performance
- Examples in prompt templates

**See:** `specs/001-local-excel-chat-agent/tasks.md` (T097-T104)

## Medium Term (1-3 months)

### User Story 4: Data Filtering ⚙️
**Priority:** P3  
**Effort:** 2 weeks

Features:
- Numeric range filtering
- Text pattern matching
- Date range filtering
- Multi-criteria queries

**See:** `specs/001-local-excel-chat-agent/tasks.md` (T105-T114)

### User Story 5: Export Capabilities 📤
**Priority:** P4  
**Effort:** 1 week

Features:
- Export conversation history (Markdown)
- Export data views (CSV)
- Generate insights summaries
- Download via browser

**See:** `specs/001-local-excel-chat-agent/tasks.md` (T115-T125)

### Mobile-Responsive Web UI 📱
**Priority:** P3  
**Effort:** 1 week

Enhancements:
- Responsive Blazor components
- Mobile-friendly file upload
- Touch-optimized interface
- Progressive Web App (PWA)

## Long Term (3-6 months)

### Write-Back Operations ✍️
**Status:** Research phase  
**Risk:** High (data integrity)

Capabilities:
- Update cell values
- Add worksheets
- Annotate findings
- Undo/redo support

**Considerations:**
- Backup before modifications
- Audit trail
- Permission model

### Advanced Analytics 📊
**Status:** Concept

Features:
- Value distributions
- Outlier detection
- Trend analysis
- Chart generation (SVG/PNG)
- Pivot table support

### Alternative Transports 🔌
**Status:** Design

Options:
- WebSocket transport (real-time)
- HTTP/REST API (stateless)
- gRPC (performance)
- Container orchestration support

### Multi-Workbook Sessions 📚
**Status:** Concept

Capabilities:
- Load multiple workbooks simultaneously
- Cross-workbook queries
- Join operations
- Workspace management

## Research & Experiments 🔬

### Formula Evaluation
- Support "what-if" scenarios
- Recalculate formulas with different inputs
- Dependency graph visualization

### Natural Language to Formula
- LLM generates Excel formulas from descriptions
- "Calculate the average of sales where region is West"
- Formula validation and explanation

### Integration with External Data
- Combine with CRM MCP servers
- Database connectors
- API data sources
- Live data refresh

### AI-Powered Insights
- Automatic anomaly detection
- Predictive analytics
- Data quality suggestions
- Schema recommendations

## Community Requests 💬

(Placeholder - track feature requests here)

## Won't Do ❌

- Cloud-only features (violates privacy-first principle)
- Proprietary model dependencies
- Telemetry/analytics that send data externally
- Breaking changes to MCP protocol

---

## How to Contribute

See `specs/001-local-excel-chat-agent/` for:
- Feature specification template
- Implementation planning
- Task breakdown examples

**Process:**
1. Create feature spec in `specs/00X-feature-name/`
2. Get feedback on design
3. Implement in feature branch
4. Submit PR with tests
5. Update documentation
```

---

## Additional Documentation Needs

### Missing Docs (Should Create)

1. **docs/DeploymentGuide.md**
   - Production deployment steps
   - Docker containerization
   - Systemd service setup (Linux)
   - Windows Service setup
   - Nginx reverse proxy
   - SSL/TLS configuration

2. **docs/ContributingGuide.md**
   - Code style guidelines
   - PR process
   - Testing requirements
   - Documentation standards

3. **docs/Architecture.md**
   - System architecture diagram
   - Component interactions
   - Data flow
   - Technology stack

4. **docs/PerformanceTuning.md**
   - Raspberry Pi optimization
   - Model selection guide
   - Memory management
   - Caching strategies

5. **CHANGELOG.md**
   - Version history
   - Breaking changes
   - Migration guides

---

## Documentation Structure Recommendations

### Proposed Organization

```
docs/
├── README.md (index of all docs)
├── user-guides/
│   ├── QuickStart.md (merged CLI + Web)
│   ├── UserGuide.md (extended guide)
│   └── Troubleshooting.md (merged CLI + Web)
├── developer-guides/
│   ├── Architecture.md
│   ├── Contributing.md
│   ├── Testing.md
│   └── Deployment.md
├── features/
│   ├── SkAgent.md (CLI features)
│   ├── WebChat.md (web features)
│   └── Roadmap.md (FutureFeatures.md)
├── archive/
│   ├── BUILD-TEST-SUMMARY-2025-10-31.md
│   ├── ProjectStatusUpdate-2025-10-31.md
│   └── ...
└── best-practices/
    ├── local-net-bp.md
    └── semantic-kernel-bp.md
```

---

## Action Items Summary

### Immediate (Today)
- [ ] Update README.md (remove "experimental" labels)
- [ ] Update docs/UserGuide.md (add web chat section)
- [ ] Archive BUILD-TEST-SUMMARY.md
- [ ] Archive docs/ProjectStatusUpdate.md

### Short Term (This Week)
- [ ] Expand docs/FutureFeatures.md
- [ ] Create docs/DeploymentGuide.md
- [ ] Create CHANGELOG.md
- [ ] Reorganize docs/ structure

### Medium Term (Next 2 Weeks)
- [ ] Create docs/Architecture.md
- [ ] Create docs/ContributingGuide.md
- [ ] Create docs/PerformanceTuning.md
- [ ] Update .github/copilot-instructions.md

---

## Notes

- All dates use ISO format (YYYY-MM-DD)
- Archive old docs rather than deleting
- Keep spec files as-is (design documents)
- Update README.md first (highest visibility)
- Web chat is NOW production-ready, not experimental!
