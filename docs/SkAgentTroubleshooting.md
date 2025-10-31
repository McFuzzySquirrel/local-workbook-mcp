# Troubleshooting the SK Agent

## Common Issues

### 1. NullReferenceException from OpenAI.Chat.ChatCompletion.get_Refusal()

**Symptom:**
```
Unhandled exception. System.NullReferenceException: Object reference not set to an instance of an object.
   at OpenAI.Chat.ChatCompletion.get_Refusal()
```

**Cause:** Version mismatch between the OpenAI SDK and your LLM server response format.

**Solutions:**

1. **Update LM Studio** to the latest version (recommended)
   - Download from https://lmstudio.ai/
   - Newer versions have better OpenAI API compatibility

2. **Check your LM Studio server settings:**
   - Make sure "Enable CORS" is checked
   - Verify the server is running on `http://localhost:1234`
   - Try loading a different model (phi-4, llama-3, etc.)

3. **Use environment variables for explicit configuration:**
   ```pwsh
   $env:LLM_BASE_URL = "http://localhost:1234/v1"
   $env:LLM_MODEL_ID = "local-model"  # Use whatever shows in LM Studio
   $env:LLM_API_KEY = "lm-studio"
   ```

4. **Try Ollama instead of LM Studio:**
   ```pwsh
   # Install Ollama from https://ollama.ai/
   ollama serve
   ollama run llama3.2
   
   # Configure the agent:
   $env:LLM_BASE_URL = "http://localhost:11434/v1"
   $env:LLM_MODEL_ID = "llama3.2"
   ```

### 2. Connection Refused or Timeout

**Symptom:** Agent hangs or shows connection errors

**Solutions:**
- Make sure your LLM server is running before starting the agent
- Check the URL - LM Studio uses `http://localhost:1234`, Ollama uses `http://localhost:11434`
- Verify your firewall isn't blocking the connection
- Try `curl http://localhost:1234/v1/models` to test the endpoint

### 3. Model Not Found

**Symptom:** Error about model not being available

**Solutions:**
- In LM Studio, make sure you've loaded a model (it should show in the server tab)
- Set `LLM_MODEL_ID` to match the exact model name shown in your LLM server
- For LM Studio, you can often use `local-model` as a generic identifier

### 4. Tools Not Being Called

**Symptom:** Agent responds but doesn't use the Excel functions

**Solutions:**
- Some models are better at function calling than others
- Try these models (in order of recommendation):
  1. `gpt-4` or `gpt-4-turbo` (if using OpenAI API)
  2. `phi-4` (good local option)
  3. `llama-3.1` or `llama-3.2` (Ollama)
  4. `mistral` (decent function calling support)
- Check your LLM server logs for errors
- Increase temperature if the model is too conservative: set in `ExcelAgent.cs`

### 5. Slow Response Times

**Solutions:**
- Use a smaller/faster model
- Check your GPU/CPU usage - make sure the LLM is using hardware acceleration
- Reduce `MaxTokens` in `ExcelAgent.cs` (currently 2000)
- Close other applications using GPU/CPU

## Testing Your LLM Server

Before running the agent, test your LLM server:

```pwsh
# Test if the server is responding
curl http://localhost:1234/v1/models

# Test a chat completion
curl http://localhost:1234/v1/chat/completions `
  -H "Content-Type: application/json" `
  -d '{
    "model": "local-model",
    "messages": [{"role": "user", "content": "Hello!"}]
  }'
```

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
