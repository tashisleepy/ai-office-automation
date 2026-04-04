# Quick Start — Copy Paste Commands

## Setup (one time)

```bash
npm install pptxgenjs docx
```

## Generate a PowerPoint Deck

### JavaScript (pptxgenjs)
```bash
node --input-type=module < scripts/create-deck.mjs
```

### Python (python-pptx) — template-based, data-driven
```bash
python3 scripts/pptx-template.py
python3 scripts/pptx-template.py --output my-deck.pptx
```

## Generate a Word Report

### JavaScript (docx-js)
```bash
node scripts/create-report.cjs
```

### Python (python-docx) — template-based, data-driven
```bash
python3 scripts/docx-template.py
python3 scripts/docx-template.py --output my-report.docx
```

## Install Ollama + Qwen (Local AI)

```bash
brew install ollama
brew services start ollama
ollama pull qwen3-coder:30b
ollama pull qwen3.5:9b
```

## Run Local AI

```bash
ollama run qwen3-coder:30b
```

## Install Instrumenta (PowerPoint Toolbar)

```bash
mkdir -p ~/Downloads/Instrumenta && \
curl -L -o ~/Downloads/Instrumenta/InstrumentaPowerpointToolbar.ppam \
  https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaPowerpointToolbar.ppam && \
curl -L -o ~/Downloads/Instrumenta/InstrumentaAppleScriptPlugin.applescript \
  https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaAppleScriptPlugin.applescript

mkdir -p ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/ && \
cp ~/Downloads/Instrumenta/InstrumentaPowerpointToolbar.ppam \
  ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/

mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint/ && \
cp ~/Downloads/Instrumenta/InstrumentaAppleScriptPlugin.applescript \
  ~/Library/Application\ Scripts/com.microsoft.Powerpoint/
```

Then open PowerPoint → Tools → PowerPoint Add-ins → enable Instrumenta.

## Install oh-my-claudecode (32 AI Agents)

```bash
npm i -g oh-my-claude-sisyphus@latest
omc setup
```

## Everything at Once

Full setup from zero:

```bash
# 1. Document generators
npm install pptxgenjs docx

# 2. Local AI
brew install ollama
brew services start ollama
ollama pull qwen3-coder:30b
ollama pull qwen3.5:9b

# 3. Instrumenta
mkdir -p ~/Downloads/Instrumenta && \
curl -L -o ~/Downloads/Instrumenta/InstrumentaPowerpointToolbar.ppam \
  https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaPowerpointToolbar.ppam && \
curl -L -o ~/Downloads/Instrumenta/InstrumentaAppleScriptPlugin.applescript \
  https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaAppleScriptPlugin.applescript && \
mkdir -p ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/ && \
cp ~/Downloads/Instrumenta/InstrumentaPowerpointToolbar.ppam \
  ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/ && \
mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint/ && \
cp ~/Downloads/Instrumenta/InstrumentaAppleScriptPlugin.applescript \
  ~/Library/Application\ Scripts/com.microsoft.Powerpoint/

# 4. clean-slides (McKinsey-style YAML → PPTX)
git clone https://github.com/tmustier/clean-slides.git /tmp/clean-slides && \
cd /tmp/clean-slides && pip3 install --break-system-packages -e . && cd - && \
pptx init

# 5. oh-my-claudecode
npm i -g oh-my-claude-sisyphus@latest
omc setup

echo "Done. Everything installed."
```

## Generate a McKinsey Slide from YAML

```bash
cat > slide.yaml << 'EOF'
title: Market Overview
subtitle: Key competitors and positioning

table:
  rows: 3
  cols: 3
  has_col_header: true
  has_row_header: true
  col_headers: ["Revenue", "Strength", "Weakness"]
  col_header_color: accent1
  row_headers:
    - "Company A"
    - "Company B"
    - "Company C"
  cells:
    - ["$50M", "Market leader", "Slow iteration"]
    - ["$12M", "Strong tech", "Limited sales"]
    - ["$3M", "Low cost", "Early stage"]
EOF

pptx generate slide.yaml -o output.pptx
open output.pptx
```
