# Future Features & Roadmap

**Last Updated:** November 1, 2025  
**Project Status:** Work in progress - CLI stable, web chat needs validation

---

## Completed ✅

- ✅ MCP Server (stdio transport)
- ✅ CLI Chat Agent (Semantic Kernel)  
- ✅ Web Chat UI (Blazor Server)
- ✅ Workbook-agnostic prompts
- ✅ Debug logging
- ✅ Linux/Raspberry Pi support
- ✅ Sample workbooks
- ✅ HTML table rendering fix (Nov 1)

---

## In Progress 🚧

**User Story 1 Validation (Phase 3)**
- Manual testing T078-T084
- End-to-end workflow validation
- Performance benchmarking

**See:** [WEB-CHAT-ROADMAP.md](../WEB-CHAT-ROADMAP.md)

---

## Short Term (Next 2-4 Weeks)

### User Story 2: Multi-Turn Conversations
**Priority:** P2 | **Effort:** 1 week
- Pronoun resolution
- Conversation summarization
- Better context preservation

**Tasks:** T085-T096

### User Story 3: Cross-Sheet Insights
**Priority:** P2 | **Effort:** 1 week
- Multi-sheet correlation
- Cross-sheet search grouping
- Performance caching

**Tasks:** T097-T104

---

## Medium Term (1-3 Months)

### User Story 4: Data Filtering (P3)
- Numeric range filtering
- Text pattern matching
- Date range filtering
- Multi-criteria queries

### User Story 5: Export Capabilities (P4)
- Export conversation history (Markdown)
- Export data views (CSV)
- Generate insights summaries

### Mobile-Responsive Web UI (P3)
- Responsive Blazor components
- Touch-optimized interface
- Progressive Web App

---

## Long Term (3-6 Months)

### Write-Back Operations
- Update cell values
- Add worksheets
- Annotate findings
- Audit trail

### Advanced Analytics
- Value distributions
- Outlier detection
- Chart generation
- Pivot table support

### Alternative Transports
- WebSocket (real-time)
- HTTP/REST (stateless)
- gRPC (performance)

---

## Research & Experiments

- Formula evaluation ("what-if" scenarios)
- Natural language to Excel formulas
- Integration with external data sources
- AI-powered insights

---

## Won't Do ❌

- Cloud-only features
- Proprietary model dependencies
- External telemetry
- Non-local-first features

---

**See:** `specs/001-local-excel-chat-agent/tasks.md` for detailed breakdown
