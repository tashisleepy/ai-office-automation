#!/bin/bash
# Ollama + Qwen Local AI Setup
# Runs entirely offline on your machine. Zero cost. Zero data sent anywhere.
#
# Requirements: macOS, 16GB+ RAM (64GB+ recommended for 30B models)
# Tested on: Mac, macOS 15+

set -e

echo "=== Installing Ollama ==="
if ! command -v ollama &> /dev/null; then
    brew install ollama
    echo "Ollama installed."
else
    echo "Ollama already installed: $(ollama --version)"
fi

echo ""
echo "=== Starting Ollama service ==="
brew services start ollama 2>/dev/null || echo "Already running"
sleep 3

echo ""
echo "=== Pulling Qwen3 Coder 30B (primary coding model) ==="
echo "This is ~18GB — will take a few minutes on first download."
ollama pull qwen3-coder:30b

echo ""
echo "=== Pulling Qwen3.5 9B (fast general model) ==="
echo "This is ~6.6GB."
ollama pull qwen3.5:9b

echo ""
echo "=== Verifying models ==="
ollama list

echo ""
echo "=== Network security check ==="
echo "Ollama should only listen on localhost:"
lsof -i -P | grep ollama | grep LISTEN

echo ""
echo "=== Setup complete! ==="
echo ""
echo "Usage:"
echo "  ollama run qwen3-coder:30b    # coding tasks (primary)"
echo "  ollama run qwen3.5:9b         # fast general tasks"
echo ""
echo "API endpoint: http://localhost:11434"
echo "Models are stored in: ~/.ollama/models/"
echo ""
echo "Both models run 100% locally. No data leaves your machine."
