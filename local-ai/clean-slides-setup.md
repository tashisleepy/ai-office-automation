# clean-slides Setup Guide

McKinsey-style PowerPoint slides from YAML. Built by an ex-McKinsey consultant. Produces minimal, information-dense tables — no graphics, no effects, just structured thinking.

## Install

```bash
git clone https://github.com/tmustier/clean-slides.git /tmp/clean-slides && \
cd /tmp/clean-slides && \
pip3 install --break-system-packages -e . && \
cd - && \
pptx init
```

## Verify

```bash
pptx --help
```

## Generate a Slide

### 1. Create a YAML spec

```yaml
# example.yaml
title: Market Overview — Key Competitors
subtitle: Funding, capabilities, and gaps

table:
  rows: 3
  cols: 3
  has_col_header: true
  has_row_header: true
  col_headers: ["Revenue", "Key Strength", "Key Weakness"]
  col_header_color: accent1

  row_headers:
    - "Company A"
    - "Company B"
    - "Company C"

  cells:
    - ["$50M ARR", "Market leader in segment", "Slow product iteration"]
    - ["$12M ARR", "Strong tech, fast shipping", "Limited enterprise sales"]
    - ["$3M ARR", "India-native, low cost", "Early stage, unproven at scale"]
```

### 2. Generate

```bash
pptx generate example.yaml -o output.pptx
```

### 3. Open

```bash
open output.pptx
```

## Features

| Feature | Syntax |
|---------|--------|
| **Bold text** | `**bold**` in any cell |
| **Italic** | `*italic*` in any cell |
| **Links** | `[text](url)` in any cell |
| **Bullet hierarchy** | Nested paragraphs with indent levels |
| **Traffic lights** | Icon indicators for RAG status |
| **Sidebars** | Split layouts (2/3, 3/4) with text |
| **Charts** | Bar, stacked, waterfall from JSON |
| **Custom templates** | Use any .pptx as base template |

## Use Your Own Template

```bash
pptx init -t your-template.pptx
```

Then edit `.clean-slides/config.yaml` to map colors and fonts.

## Common Commands

```bash
# Generate slide from YAML
pptx generate spec.yaml -o output.pptx

# List slides in a deck
pptx list deck.pptx

# Summarize a slide
pptx summary deck.pptx 1

# Inspect theme colors
pptx theme deck.pptx

# Merge multiple decks
pptx merge deck1.pptx deck2.pptx -o combined.pptx

# Edit text in an existing slide
pptx edit deck.pptx 1 shape_idx "New text"
```

## Links

- GitHub: https://github.com/tmustier/clean-slides
- Design philosophy: https://mustier.ai/projects/clean-slides
- License: MIT
