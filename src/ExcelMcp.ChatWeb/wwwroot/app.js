const chatLog = document.getElementById("chat-log");
const chatForm = document.getElementById("chat-form");
const chatText = document.getElementById("chat-text");
const sendButton = document.getElementById("send-button");
const yearLabel = document.getElementById("year");
const modelLabel = document.getElementById("model-label");
const toolsPanel = document.getElementById("tools-panel");
const previewToggle = document.getElementById("preview-toggle");
const previewForm = document.getElementById("preview-form");
const previewWorksheetInput = document.getElementById("preview-worksheet");
const previewTableInput = document.getElementById("preview-table");
const previewRowsInput = document.getElementById("preview-rows");
const previewStatus = document.getElementById("preview-status");
const previewOutput = document.getElementById("preview-output");
const previewLoadMoreButton = document.getElementById("preview-load-more");
const worksheetOptions = document.getElementById("worksheet-options");

const conversation = [];
let isBusy = false;

const previewState = {
    worksheet: "",
    table: "",
    rows: 10,
    nextCursor: null,
    hasMore: false,
    tableElement: null,
    worksheetOptionsLoaded: false
};

function updateYear() {
    const now = new Date();
    if (yearLabel) {
        yearLabel.textContent = String(now.getFullYear());
    }
}

function setBusy(state) {
    isBusy = state;
    if (sendButton) {
        sendButton.disabled = state;
        sendButton.textContent = state ? "Thinking…" : "Send";
    }
    if (chatText) {
        chatText.disabled = state;
    }
}

function sanitize(text) {
    return text.replace(/[\u0000-\u001f\u007f]/g, "");
}

function createBubble(role, content) {
    const wrapper = document.createElement("div");
    wrapper.classList.add("chat-bubble", role);

    const roleLabel = document.createElement("div");
    roleLabel.className = "role";
    roleLabel.textContent = roleLabelFor(role);

    const message = document.createElement("div");
    message.className = "content";
    message.textContent = content;

    wrapper.append(roleLabel, message);
    return wrapper;
}

function roleLabelFor(role) {
    switch (role) {
        case "assistant":
            return "Workbook Copilot";
        case "user":
            return "You";
        case "tool":
            return "Tool";
        default:
            return role.charAt(0).toUpperCase() + role.slice(1);
    }
}

function appendMessage(role, content, { persist = true } = {}) {
    const text = sanitize(content);
    const bubble = createBubble(role, text);
    chatLog?.appendChild(bubble);
    chatLog?.scrollTo({ top: chatLog.scrollHeight, behavior: "smooth" });

    if (persist && (role === "assistant" || role === "user")) {
        conversation.push({ role, content: text });
    }
}

async function sendMessage(prompt) {
    if (!prompt.trim()) {
        return;
    }

    appendMessage("user", prompt);
    setBusy(true);

    try {
        const payload = { messages: conversation.slice() };
        const response = await fetch("/api/chat", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            throw new Error(`Server responded with ${response.status}`);
        }

        const result = await response.json();
        if (Array.isArray(result.toolCalls)) {
            for (const tool of result.toolCalls) {
                const args = tool.arguments ? JSON.stringify(tool.arguments) : "{}";
                const summary = tool.outputSummary ?? "(no summary)";
                const tag = tool.isError ? "⚠ Tool error" : "Tool call";
                appendMessage("tool", `${tag}: ${tool.name}\nInput: ${args}\n${summary}`, { persist: false });
            }
        }

        if (typeof result.reply === "string" && result.reply.trim().length > 0) {
            appendMessage("assistant", result.reply.trim());
        } else {
            appendMessage("assistant", "I did not receive a reply." );
        }
    } catch (error) {
        const reason = error instanceof Error ? error.message : String(error);
        appendMessage("assistant", `I ran into a problem: ${reason}`, { persist: false });
    } finally {
        setBusy(false);
    }
}

async function refreshModelInfo() {
    if (!modelLabel) {
        return;
    }

    try {
        const response = await fetch("/api/model", { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }
        const data = await response.json();
        const model = typeof data.model === "string" && data.model.trim().length > 0 ? data.model.trim() : "unknown";
        const baseUrl = typeof data.baseUrl === "string" && data.baseUrl.trim().length > 0 ? data.baseUrl.trim() : "";
        modelLabel.textContent = baseUrl ? `Model: ${model} @ ${baseUrl}` : `Model: ${model}`;
    } catch (error) {
        const reason = error instanceof Error ? error.message : String(error);
        modelLabel.textContent = `Model: unavailable (${reason})`;
    }
}

function setPreviewStatus(message, isError = false) {
    if (!previewStatus) {
        return;
    }

    previewStatus.textContent = message ?? "";
    previewStatus.classList.toggle("error", Boolean(isError));
}

function resetPreviewOutput() {
    if (previewOutput) {
        previewOutput.innerHTML = "";
    }
    previewState.tableElement = null;
}

function ensurePreviewTable(headers) {
    if (!previewOutput) {
        return null;
    }

    if (previewState.tableElement) {
        return previewState.tableElement;
    }

    const table = document.createElement("table");
    table.className = "preview-table";

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    const headerCells = ["Row #", ...(headers ?? [])];

    for (const label of headerCells) {
        const th = document.createElement("th");
        th.textContent = sanitize(String(label ?? ""));
        headerRow.appendChild(th);
    }

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    table.appendChild(tbody);

    previewOutput.innerHTML = "";
    previewOutput.appendChild(table);
    previewState.tableElement = table;
    return table;
}

function appendPreviewRows(rows) {
    if (!previewState.tableElement || !rows || rows.length === 0) {
        return;
    }

    const tbody = previewState.tableElement.tBodies[0] ?? previewState.tableElement.createTBody();
    const headerCellCount = previewState.tableElement.tHead?.rows?.[0]?.cells?.length ?? 1;
    const dataCellCount = Math.max(0, headerCellCount - 1);

    for (const row of rows) {
        const tr = document.createElement("tr");

        const rowNumberCell = document.createElement("td");
    const rowNumberValue = row.rowNumber ?? row.RowNumber ?? "";
    rowNumberCell.textContent = String(rowNumberValue);
        tr.appendChild(rowNumberCell);

    const values = Array.isArray(row.values ?? row.Values) ? (row.values ?? row.Values) : [];
        for (let index = 0; index < dataCellCount; index += 1) {
            const value = index < values.length ? values[index] : "";
            const td = document.createElement("td");
            td.textContent = sanitize(value == null ? "" : String(value));
            tr.appendChild(td);
        }

        tbody.appendChild(tr);
    }
}

function updateWorksheetOptions(worksheets) {
    if (!worksheetOptions) {
        return;
    }

    worksheetOptions.innerHTML = "";
    if (!Array.isArray(worksheets) || worksheets.length === 0) {
        return;
    }

    for (const name of worksheets) {
        const option = document.createElement("option");
        option.value = name;
        worksheetOptions.appendChild(option);
    }
}

async function refreshWorksheetOptions() {
    if (previewState.worksheetOptionsLoaded || !worksheetOptions) {
        return;
    }

    try {
        const response = await fetch("/api/resources", { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        const worksheets = Array.isArray(data?.worksheets) ? data.worksheets : [];
        updateWorksheetOptions(worksheets);
        previewState.worksheetOptionsLoaded = true;
    } catch (error) {
        // Surface in status if panel is visible
        const message = error instanceof Error ? error.message : String(error);
        setPreviewStatus(`Unable to load worksheet names: ${message}`, true);
    }
}

async function loadPreview(options) {
    if (!previewLoadMoreButton) {
        return;
    }

    previewLoadMoreButton.disabled = true;
    setPreviewStatus("Loading preview…", false);

    try {
        const payload = {
            worksheet: options.worksheet,
            rows: options.rows
        };

        if (options.table) {
            payload.table = options.table;
        }

        if (options.cursor) {
            payload.cursor = options.cursor;
        }

        const response = await fetch("/api/preview", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(payload)
        });

        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await response.json();
        const success = data?.success ?? data?.Success ?? false;
        if (!success) {
            resetPreviewOutput();
            setPreviewStatus(data?.error ?? data?.Error ?? "Preview failed.", true);
            return;
        }

        if (!options.append) {
            resetPreviewOutput();
        }

        const headers = Array.isArray(data?.headers)
            ? data.headers
            : Array.isArray(data?.Headers)
                ? data.Headers
                : [];

        const rows = Array.isArray(data?.rows)
            ? data.rows
            : Array.isArray(data?.Rows)
                ? data.Rows
                : [];

        const table = ensurePreviewTable(headers);
        if (table) {
            appendPreviewRows(rows);
        }

        const nextCursor = data?.nextCursor ?? data?.NextCursor ?? null;
        previewState.nextCursor = typeof nextCursor === "string" && nextCursor.trim().length > 0 ? nextCursor.trim() : null;
        const hasMore = data?.hasMore ?? data?.HasMore ?? false;
        previewState.hasMore = Boolean(hasMore && previewState.nextCursor);

        const offset = data?.offset ?? data?.Offset ?? 0;

        if (previewState.hasMore && previewState.nextCursor) {
            setPreviewStatus(`Showing rows from offset ${offset}. Next cursor: ${previewState.nextCursor}.`, false);
            previewLoadMoreButton.disabled = false;
        } else {
            setPreviewStatus(`Showing rows from offset ${offset}. No additional pages.`, false);
        }
    } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        setPreviewStatus(`Preview failed: ${message}`, true);
        if (!options.append) {
            resetPreviewOutput();
        }
    }
}

chatForm?.addEventListener("submit", async event => {
    event.preventDefault();
    if (isBusy) {
        return;
    }

    const text = chatText?.value ?? "";
    const cleaned = text.trim();
    if (!cleaned) {
        return;
    }

    chatText.value = "";
    await sendMessage(cleaned);
    chatText.focus();
});

chatText?.addEventListener("keydown", event => {
    if (event.key === "Enter" && !event.shiftKey) {
        event.preventDefault();
        chatForm?.requestSubmit();
    }
});

previewForm?.addEventListener("submit", async event => {
    event.preventDefault();
    const worksheet = previewWorksheetInput?.value?.trim();
    if (!worksheet) {
        setPreviewStatus("Worksheet is required.", true);
        previewWorksheetInput?.focus();
        return;
    }

    if (toolsPanel?.classList.contains("collapsed")) {
        togglePreview(true);
    }

    const table = previewTableInput?.value?.trim() ?? "";
    const rowsValue = Number.parseInt(previewRowsInput?.value ?? "", 10);
    const rows = Number.isFinite(rowsValue) && rowsValue > 0 ? Math.min(rowsValue, 100) : 10;

    previewState.worksheet = worksheet;
    previewState.table = table;
    previewState.rows = rows;
    previewState.nextCursor = null;
    previewState.hasMore = false;

    await loadPreview({ worksheet, table: table || undefined, rows, cursor: null, append: false });
});

previewLoadMoreButton?.addEventListener("click", async () => {
    if (!previewState.hasMore || !previewState.nextCursor) {
        return;
    }

    if (toolsPanel?.classList.contains("collapsed")) {
        togglePreview(true);
    }

    await loadPreview({
        worksheet: previewState.worksheet,
        table: previewState.table || undefined,
        rows: previewState.rows,
        cursor: previewState.nextCursor,
        append: true
    });
});

function togglePreview(forceExpand) {
    if (!toolsPanel || !previewToggle) {
        return;
    }

    const shouldExpand = typeof forceExpand === "boolean" ? forceExpand : toolsPanel.classList.contains("collapsed");

    if (shouldExpand) {
        toolsPanel.classList.remove("collapsed");
        previewToggle.textContent = "Hide";
        previewToggle.setAttribute("aria-expanded", "true");
        refreshWorksheetOptions();
    } else {
        toolsPanel.classList.add("collapsed");
        previewToggle.textContent = "Show";
        previewToggle.setAttribute("aria-expanded", "false");
    }
}

previewToggle?.addEventListener("click", () => {
    const willExpand = toolsPanel?.classList.contains("collapsed") ?? true;
    togglePreview(willExpand);
});

appendMessage("assistant", "Hello! Ask me anything about your workbook.");
updateYear();
refreshModelInfo();
refreshWorksheetOptions();
chatText?.focus();
