# Local AI Configuration

## Installed Models

| Model | Size | Use Case | Speed |
|-------|------|----------|-------|
| **qwen3-coder:30b** | 18 GB | Coding, refactoring, debugging, multi-file changes | ~3-5 sec/response |
| **qwen3.5:9b** | 6.6 GB | General tasks, multimodal, tool calling | ~5-10 sec/response |

## Quick Commands

```bash
# Interactive coding session
ollama run qwen3-coder:30b

# API call (for scripts/integrations)
curl -s http://localhost:11434/api/generate \
  -d '{"model":"qwen3-coder:30b","prompt":"Your prompt here","stream":false}'

# Check loaded models
curl -s http://localhost:11434/api/ps

# List all models
ollama list
```

## Security

- Ollama listens on **localhost:11434 only** — not exposed to network
- Model files are GGUF format — inert weight data, no executable code
- Models pulled from official Ollama library (registry.ollama.ai)
- Publisher: Alibaba Cloud (Qwen team), Apache 2.0 license
- SHA256 digests verified on download

## Hardware Requirements

| Model | Min RAM | Recommended RAM |
|-------|---------|-----------------|
| qwen3-coder:30b | 20 GB | 32 GB+ |
| qwen3.5:9b | 8 GB | 16 GB+ |
| Both simultaneously | 28 GB | 64 GB+ |

## Service Management

```bash
# Start on boot
brew services start ollama

# Stop
brew services stop ollama

# Update models
ollama pull qwen3-coder:30b
ollama pull qwen3.5:9b
```
