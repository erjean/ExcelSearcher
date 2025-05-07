# Excel Translator Search Tool

A lightweight and fast desktop app that quickly searches through multiple Excel files and locate a specified string.

I created this tool to support my own translation workflow. Since I frequently work with Excel files, I wanted something lightweight and efficient. Current free solutions did not seem to provide what I needed.

---

## Features

- Search through multiple `.xlsx` files in a folder
- Displays matches grouped by file and sheet
- Recent folders and favorite folders for fast access
- Auto-adjusting column widths based on actual content
- Drag to scroll both vertically and horizontally
- Toggle options for partial match and case sensitivity
- Zoom in/out for font scaling
- Lightweight - no dependencies beyond `openpyxl`
- Packaged as a standalone Windows `.exe` (no Python install needed)

---
## Getting Started

### Option 1: Use the EXE

1. Download `excelsearcher.exe` from the [Releases section](https://github.com/erjean/ExcelSearcher/releases) or [download directly](https://github.com/erjean/ExcelSearcher/releases/download/v1.0.0/Excel.Searcher.exe)
2. Double-click and run - no install needed

### Option 2: Run the Python script

```bash
pip install -r requirements.txt
python ExcelSearcher.py
```

---

## Requirements

- Python 3.6+
- `openpyxl`
