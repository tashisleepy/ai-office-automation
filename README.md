# AI Office Automation

Programmatic generation of enterprise-grade PowerPoint decks and Word documents, plus a fully local AI coding setup. No subscriptions. No cloud dependency. Production-ready.

## What This Does

| Capability | Tool | Output |
|-----------|------|--------|
| **Deck generation** | pptxgenjs | .pptx files with consulting-grade layouts, branded colors, numbered sections |
| **Report generation** | docx-js | .docx files with headers, footers, styled tables, corporate branding |
| **Local AI** | Ollama + Qwen | 30B coding model running 100% offline on your machine |

## Quick Start

### Generate a PowerPoint Deck
```bash
npm install pptxgenjs
node --input-type=module < scripts/create-deck.mjs
# Output: output-deck.pptx
```

### Generate a Word Report
```bash
npm install docx
node scripts/create-report.cjs
# Output: output-report.docx
```

### Set Up Local AI
```bash
chmod +x local-ai/setup.sh
./local-ai/setup.sh
# Installs Ollama + pulls Qwen3 Coder 30B + Qwen3.5 9B
```

## Document Generation

Both generators follow consulting-firm standards:
- **McKinsey-style layouts** — clean, minimal, data-heavy
- **Consistent branding** — edit the `BRAND` config object once, applies everywhere
- **Programmatic** — feed data from any source (APIs, databases, AI analysis) into templates
- **No PowerPoint/Word required** to generate — only to view the output

### Deck Features
- Widescreen layout (16:9)
- Dark title/contact slides with light content slides
- Numbered section items with colored circles
- Accent divider lines
- Brand-configurable colors and fonts

### Report Features
- Professional headers with brand line and confidentiality notice
- Page numbers in footer
- Auto-styled tables with alternating row shading
- Navy header rows with white text
- Heading hierarchy (H1/H2) with proper spacing

## Local AI Setup

Two models, zero cost, fully offline:

| Model | Size | Best For |
|-------|------|----------|
| **qwen3-coder:30b** | 18 GB | Coding, refactoring, debugging (PRIMARY) |
| **qwen3.5:9b** | 6.6 GB | Fast general tasks, multimodal |

### Security
- Runs on **localhost only** — not exposed to network
- GGUF format — inert weight data, no executable code
- Official Ollama library source, SHA256 verified
- Apache 2.0 licensed (Alibaba Cloud / Qwen team)

### Usage
```bash
# Interactive coding
ollama run qwen3-coder:30b

# API integration
curl -s http://localhost:11434/api/generate \
  -d '{"model":"qwen3-coder:30b","prompt":"Write a fibonacci function","stream":false}'
```

## Project Structure

```
ai-office-automation/
├── scripts/
│   ├── create-deck.mjs      # PowerPoint generator (pptxgenjs)
│   └── create-report.cjs    # Word document generator (docx-js)
├── local-ai/
│   ├── setup.sh             # One-command Ollama + Qwen install
│   └── ollama-config.md     # Model details, commands, security notes
├── templates/               # Custom brand templates (add your own)
└── docs/                    # Additional documentation
```

## Dependencies

```bash
# Document generation
npm install pptxgenjs docx

# Local AI
brew install ollama
ollama pull qwen3-coder:30b
ollama pull qwen3.5:9b
```

## Hardware Tested

- Mac, macOS 15+
- Both AI models run simultaneously with 100+ GB headroom

## License

MIT

---

*Built by Tashi. Documents generated for enterprise clients at scale.*
