const chatLog = document.getElementById("chat-log");
const chatForm = document.getElementById("chat-form");
const chatText = document.getElementById("chat-text");
const sendButton = document.getElementById("send-button");
const yearLabel = document.getElementById("year");

const conversation = [];
let isBusy = false;

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

appendMessage("assistant", "Hello! Ask me anything about your workbook.");
updateYear();
chatText?.focus();
