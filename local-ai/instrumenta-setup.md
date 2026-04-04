# Instrumenta Setup Guide

Free open-source PowerPoint toolbar with 270+ consulting-grade formatting tools. Mimics the proprietary add-ins used at McKinsey, BCG, and Bain.

## What You Get

- Perfect alignment and distribution
- Table optimization and formatting
- Harvey Balls and traffic light indicators
- RAG (Red/Amber/Green) status indicators
- Pyramid storyline builder
- Bulk formatting across slides
- Mail-merge from Excel
- Data anonymization
- Professional shape and object tools

## One-Command Install (macOS)

### Step 1: Download

```bash
mkdir -p ~/Downloads/Instrumenta && \
gh release download 1.66 --repo iappyx/Instrumenta --dir ~/Downloads/Instrumenta
```

If you don't have `gh` CLI:

```bash
mkdir -p ~/Downloads/Instrumenta && \
curl -L -o ~/Downloads/Instrumenta/InstrumentaPowerpointToolbar.ppam \
  https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaPowerpointToolbar.ppam && \
curl -L -o ~/Downloads/Instrumenta/InstrumentaAppleScriptPlugin.applescript \
  https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaAppleScriptPlugin.applescript
```

### Step 2: Install Add-in

```bash
mkdir -p ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/ && \
cp ~/Downloads/Instrumenta/InstrumentaPowerpointToolbar.ppam \
  ~/Library/Group\ Containers/UBF8T346G9.Office/User\ Content/Add-Ins/
```

### Step 3: Install AppleScript Plugin (for export features)

```bash
mkdir -p ~/Library/Application\ Scripts/com.microsoft.Powerpoint/ && \
cp ~/Downloads/Instrumenta/InstrumentaAppleScriptPlugin.applescript \
  ~/Library/Application\ Scripts/com.microsoft.Powerpoint/
```

### Step 4: Activate in PowerPoint

1. Open PowerPoint
2. Go to **Tools → PowerPoint Add-ins**
3. Click **+** (plus) button
4. Navigate to the Add-Ins folder and select `InstrumentaPowerpointToolbar.ppam`
5. Click **OK**
6. You'll see a new **Instrumenta** tab in the ribbon

## One-Command Install (Windows)

### Download

```powershell
mkdir ~\Downloads\Instrumenta
Invoke-WebRequest -Uri "https://github.com/iappyx/Instrumenta/releases/download/1.66/InstrumentaPowerpointToolbar.ppam" -OutFile ~\Downloads\Instrumenta\InstrumentaPowerpointToolbar.ppam
```

### Install

1. Double-click `InstrumentaPowerpointToolbar.ppam`
2. Click **Enable Macros** when prompted
3. Instrumenta tab appears in the ribbon

## Verify Installation

Open PowerPoint. You should see the **Instrumenta** tab between the existing ribbon tabs. Click it — you'll see sections for:
- **Align** — alignment and distribution tools
- **Format** — bulk formatting, styles
- **Tables** — table optimization
- **Shapes** — professional shape tools
- **Slides** — slide management
- **Tools** — Harvey Balls, traffic lights, anonymizer

## Requirements

- Microsoft PowerPoint (Microsoft 365 or Office 2016+)
- macOS or Windows

## Links

- GitHub: https://github.com/iappyx/Instrumenta
- Latest release: https://github.com/iappyx/Instrumenta/releases
- License: MIT (free forever)
