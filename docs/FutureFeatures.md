# Future Features & Roadmap

**Last Updated:** March 12, 2026  
**Project Status:** MVP Complete — MCP SDK migration, write operations, and pivot analysis all shipped

---

## Completed ✅

- ✅ MCP Server (stdio transport, ModelContextProtocol SDK 1.1.0)
- ✅ CLI Chat Agent (Semantic Kernel 1.73.0)
- ✅ Web Chat UI (Blazor Server)
- ✅ Workbook-agnostic prompts
- ✅ Debug logging (Serilog structured)
- ✅ Linux/Raspberry Pi support
- ✅ Sample workbooks + pivot test workbook
- ✅ HTML table rendering
- ✅ User Story 1: Basic Workbook Querying
- ✅ User Story 2: Multi-Turn Conversations
- ✅ User Story 3: Cross-Sheet Insights
- ✅ User Story 4: Data Filtering
- ✅ User Story 5: Export Capabilities (CSV + Markdown)
- ✅ **Write-Back Operations** — `excel-write-cell`, `excel-write-range`, `excel-create-worksheet` with timestamped auto-backup
- ✅ **Pivot Table Analysis** — `excel-analyze-pivot` with field extraction and aggregated data
- ✅ **Official MCP SDK 1.1.0** — full spec compliance; `[McpServerTool]` attribute-based registration
- ✅ **External MCP Client Configs** — Claude Desktop, GitHub Copilot, Cursor out-of-box
- ✅ **Ollama as default LLM** — auto-detection with LM Studio fallback
- ✅ Real-time token streaming in Blazor UI
- ✅ Session export to CSV/Markdown

---

## In Progress 🚧

**Phase 5+: Polish & Extensibility**
- Unit test expansion for write operations
- Provider status indicator in Chat UI
- Stream cancel button in UI

---

## Short Term (Next 2-4 Weeks)

### Mobile-Responsive Web UI
- Responsive Blazor components
- Touch-optimized interface

---

## Medium Term (1-3 Months)

### Pivot Table Enhancements
- Pivot table creation/modification via MCP tool
- Drill-down from pivot cells to source data rows
- Pivot cache and calculated field analysis
- Slicer and filter integration

### Formula Support
- Formula evaluation ("what-if" scenarios)
- Natural language to Excel formula translation

---

## Long Term (3-6 Months)

### Alternative Transports
- WebSocket (real-time, bi-directional)
- HTTP/REST (stateless API mode)

### Advanced Analytics
- Value distribution and outlier detection
- Chart generation via OpenXML

---

## Research & Experiments

- Formula evaluation ("what-if" scenarios)
- Natural language to Excel formulas
- Integration with external data sources
- AI-powered anomaly detection

---

## Won't Do ❌

- Cloud-only features
- Proprietary model dependencies
- External telemetry
- Non-local-first features

---

**See:** `specs/001-local-excel-chat-agent/tasks.md` for detailed breakdown
