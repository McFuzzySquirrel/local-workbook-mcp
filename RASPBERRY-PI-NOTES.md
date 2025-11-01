# Running on Raspberry Pi

**Last Updated:** November 1, 2025

## ✅ What Works

- CLI Agent builds and runs successfully
- Web Chat builds and runs successfully  
- MCP Server works perfectly
- LM Studio server runs on ARM64
- Model detection now working

## ⚠️ Performance Considerations

### Model Size vs Speed

Your Raspberry Pi is running **Qwen3-8B** which is quite large for the Pi's CPU/RAM. Here's what we observed:

**From LM Studio logs:**
```
PromptProcessing: 46.3348 seconds (almost 5 minutes!)
```

This means the model took **~46 seconds just to process your prompt** before even starting to generate a response.

### Recommended Models for Raspberry Pi

For better performance on Raspberry Pi, use smaller models:

| Model Size | Speed | Quality | Recommendation |
|------------|-------|---------|----------------|
| **1.5B-3B** | ⚡⚡⚡ Fast | ⭐⭐ Good | ✅ **Best for Pi** |
| **3B-7B** | ⚡⚡ Medium | ⭐⭐⭐ Better | ⚠️ Usable but slow |
| **8B+** | ⚡ Slow | ⭐⭐⭐⭐ Best | ❌ Too slow for Pi |

**Recommended models for Raspberry Pi:**
- `qwen2.5-1.5b-instruct` ⭐ **BEST CHOICE**
- `phi-3-mini` (3.8B)
- `tinyllama-1.1b`
- `stablelm-zephyr-3b`

### Current Setup

```
Model: qwen/qwen3-8b
Threads: 3
Context: 4096 tokens
Prompt Processing: ~46 seconds
```

### How to Switch Models in LM Studio

1. Stop the current server
2. Load a smaller model (e.g., qwen2.5-1.5b-instruct)
3. Start the server again
4. Run the CLI - it will auto-detect the new model!

## Timeout Settings

We've increased timeouts to handle slower models:

- **HTTP Client:** 5 minutes
- **Request Timeout:** 3 minutes
- **Model Detection:** 5 seconds

This prevents "Client disconnected" errors, but doesn't make the model faster.

## CLI Features Working

✅ Model detection shows actual running model  
✅ LLM server warning if not running  
✅ Colored terminal output  
✅ Workbook switching  
✅ Debug logging  

## Performance Tips

1. **Use smaller models** (1.5B-3B range)
2. **Reduce context window** in LM Studio (try 2048 instead of 4096)
3. **Lower max tokens** if responses are slow
4. **Close other apps** to free RAM
5. **Consider quantized models** (Q4 or Q5)

## Monitoring Performance

Check LM Studio logs to see:
- Prompt processing time
- Token generation speed
- Memory usage

**Location:** `~/.lmstudio/server-logs`

## Questions?

See `docs/SkAgentQuickStart.md` for general CLI usage.

---

**Bottom line:** The 8B model works but is too slow for comfortable use on Pi. Switch to qwen2.5-1.5b-instruct for much better performance! 🚀
