# Troubleshooting the SK Agent

**Last Updated:** March 12, 2026

## Common Issues

### 1. Connection Refused or Timeout

**Symptom:** Agent hangs or shows connection errors on startup.

**Solutions:**
- Make sure your LLM server is running *before* starting the agent
- Ollama default: `http://localhost:11434` — verify with `curl http://localhost:11434/api/tags`
- LM Studio default: `http://localhost:1234` — verify with `curl http://localhost:1234/v1/models`
- Check your firewall isn't blocking the port
- Set explicit URL via `LLM_BASE_URL` (include `/v1` suffix):

```bash
export LLM_BASE_URL="http://localhost:11434/v1"   # Ollama
export LLM_BASE_URL="http://localhost:1234/v1"    # LM Studio
```

### 2. NullReferenceException from OpenAI SDK

**Symptom:**
```
System.NullReferenceException: Object reference not set to an instance of an object.
   at OpenAI.Chat.ChatCompletion.get_Refusal()
```

**Cause:** Version mismatch between the OpenAI SDK and your LLM server response format.

**Solutions:**

1. **Update LM Studio** to the latest version — newer versions have better OpenAI API compatibility
2. **Switch to Ollama** (generally more stable for function calling):
   ```bash
   ollama pull llama3.2
   ollama serve
   export LLM_BASE_URL="http://localhost:11434/v1"
   export LLM_MODEL_ID="llama3.2"
   ```
3. **Use explicit env vars for LM Studio:**
   ```bash
   export LLM_BASE_URL="http://localhost:1234/v1"
   export LLM_MODEL_ID="local-model"
   export LLM_API_KEY="lm-studio"
   ```

### 3. Model Not Found

**Symptom:** Error about model not being available.

**Solutions:**
- For Ollama: verify the model is pulled — `ollama list`
- For LM Studio: make sure a model is loaded in the server tab (not just downloaded)
- Set `LLM_MODEL_ID` to the exact model name shown in your server
- For LM Studio, `local-model` often works as a generic identifier

### 4. Tools Not Being Called

**Symptom:** Agent responds conversationally but never calls Excel functions (no data returned).

**Solutions:**

Models vary significantly in function-calling quality. Try these (best first):

| Model | Provider | Notes |
|---|---|---|
| `llama3.2` | Ollama | Best local option for tool use |
| `llama3.1` | Ollama | Solid function calling |
| `phi-4` | LM Studio / Ollama | Good reasoning + tool use |
| `mistral` | Ollama | Decent function calling |
| `gpt-4` / `gpt-4-turbo` | OpenAI API | Best overall, requires key |

Additional steps:
- Check your LLM server logs for errors during tool invocation
- Make sure the workbook path is valid and accessible (`EXCEL_MCP_WORKBOOK` or `--workbook` arg)

### 5. Slow Response Times

**Solutions:**
- Use a smaller/faster model (e.g., `llama3.2:3b` instead of `llama3.1:70b`)
- Verify the LLM is using hardware acceleration (GPU) — check server logs
- Reduce `MaxTokens` in `src/ExcelMcp.SkAgent/ExcelAgent.cs` (currently 2000)
- Close other GPU-intensive applications

### 6. Write Operation Not Creating Backup

**Symptom:** `write_cell`, `write_range`, or `create_worksheet` completes but no backup file is created.

**Solutions:**
- Verify the workbook directory is writable: `ls -la $(dirname $EXCEL_MCP_WORKBOOK)`
- Backups are written to the same directory as the workbook, named `<workbook>_<timestamp>Z.xlsx`
- Check the MCP server logs for backup-related errors
- If running from a read-only filesystem (e.g., a network share), move the workbook to a local writable path

### 7. Workbook Path Not Found

**Symptom:** `"Workbook file not found"` or `FileNotFoundException` on first tool call.

**Solutions:**
- Use an absolute path: `export EXCEL_MCP_WORKBOOK="/home/user/data/file.xlsx"`
- Verify the file exists: `ls -la "$EXCEL_MCP_WORKBOOK"`
- Generate a sample workbook for testing:
  ```bash
  pwsh scripts/create-sample-workbooks.ps1
  export EXCEL_MCP_WORKBOOK="$(pwd)/test-data/ProjectTracking.xlsx"
  ```

---

## Testing Your LLM Server

Before running the agent, verify your LLM server is healthy:

```bash
# Ollama
curl http://localhost:11434/api/tags
curl http://localhost:11434/v1/models

# LM Studio
curl http://localhost:1234/v1/models

# Test a chat completion (either server)
curl http://localhost:11434/v1/chat/completions \
  -H "Content-Type: application/json" \
  -d '{
    "model": "llama3.2",
    "messages": [{"role": "user", "content": "Say hello"}]
  }'
```

---

## Recommended Setup

**For best results:**

1. **Ollama** with one of these models:
   ```bash
   ollama pull llama3.2          # 3B — fast, good tool use
   ollama pull llama3.1          # 8B — better reasoning
   ollama pull phi4              # 14B — excellent if you have the VRAM
   ```

2. **Or LM Studio 0.3.0+** with:
   - `meta-llama/Llama-3.2-3B-Instruct`
   - `microsoft/phi-4`
   - `mistralai/Mistral-7B-Instruct-v0.3`

3. **Environment variables (Linux/macOS):**
   ```bash
   export LLM_BASE_URL="http://localhost:11434/v1"
   export LLM_MODEL_ID="llama3.2"
   ```

---

## Still Having Issues?

1. Try the **Web UI** (`ExcelMcp.ChatWeb`) — it's been more extensively tested and has better error surfacing
2. Use the **CLI debug tool** to verify raw MCP calls work before involving the LLM:
   ```bash
   dotnet run --project src/ExcelMcp.Client -- list
   dotnet run --project src/ExcelMcp.Client -- preview "SheetName"
   ```
3. File an issue on GitHub with:
   - Your Ollama/LM Studio version
   - The model you're using
   - Full error message and stack trace
   - Output of `curl http://localhost:11434/api/tags` (or LM Studio equivalent)

If these work, your server is configured correctly.

## Recommended Setup

**For best results:**

1. **LM Studio 0.3.0+** with one of these models:
   - microsoft/phi-4
   - meta-llama/Llama-3.2-3B-Instruct
   - mistralai/Mistral-7B-Instruct-v0.3

2. **Or Ollama** with:
   ```bash
   ollama pull llama3.2
   ollama serve
   ```

3. **Environment variables:**
   ```pwsh
   $env:LLM_BASE_URL = "http://localhost:1234/v1"  # Add /v1!
   $env:LLM_MODEL_ID = "local-model"
   $env:LLM_API_KEY = "lm-studio"
   ```

## Still Having Issues?

1. Check the [LM Studio Discord](https://discord.gg/aPQfnNkxGC) for known issues
2. Try the web UI version (`ExcelMcp.ChatWeb`) which has been tested more extensively
3. Use the direct MCP server + client if you just need to query data without conversational AI
4. File an issue on GitHub with:
   - Your LM Studio/Ollama version
   - The model you're using
   - Full error message
   - Output of `curl http://localhost:1234/v1/models`
