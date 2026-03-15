# PDFree

A free, open-source PDF toolbox desktop application built with Python and PySide6.

![PDFree logo](LOGO.svg)

## Features

- **View PDF** — Full-featured viewer with zoom, rotation, text selection, search (Ctrl+F), thumbnails, TOC sidebar, annotations, signature drawing, and form filling
- **Excerpt Tool** — Load multiple PDFs, drag to select rectangular regions from any page, and collect them into a new PDF
- **Split** — Split a PDF by page ranges, every N pages, or bookmarks
- **File Library** — Persistent library of your PDFs with folders, favorites, and recent files
- **PDF to CSV** — Extract tables from PDFs into CSV files

More tools (merge, crop, watermark, password, OCR, etc.) are shown in the UI and will be added in future releases.

## Installation

### Option 1 — Pre-built app (no Python required)

**Windows**
1. Go to the [Releases page](https://github.com/Fioerd/PDFree/releases)
2. Download `PDFree_Setup.exe` from the latest release
3. Run the installer

> Windows SmartScreen may show an "Unknown publisher" warning on first launch. Click **More info → Run anyway** to proceed.

**macOS**
1. Go to the [Releases page](https://github.com/Fioerd/PDFree/releases)
2. Download `PDFree.dmg` from the latest release
3. Open the `.dmg` and drag **PDFree.app** to your Applications folder

> Gatekeeper may block the app on first launch because it is not notarized. Right-click (or Control-click) the app → **Open** → **Open** to allow it.

---

### Option 2 — Run from source (Windows, macOS, Linux)

**Requirements:** Python 3.8 or newer · Tested on Windows and macOS · Linux should work but is untested

```bash
# 1. Clone the repository
git clone https://github.com/Fioerd/PDFree.git
cd PDFree
```

**Windows**
```bash
# 2. Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

**macOS — quick setup (Python 3.8+ already installed)**
```bash
# 2. Create and activate a virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

**macOS / Linux — with Homebrew**
```bash
# 2. Install Python 3.11
brew install python@3.11

# 3. Create a virtual environment with Python 3.11
python3.11 -m venv .venv
source .venv/bin/activate

# 4. Install dependencies
pip install -r requirements.txt
```

**macOS — without Homebrew**

1. Download and install Python 3.11 from the official website: https://www.python.org/downloads/
2. Open Terminal and run:

```bash
# 2. Create a virtual environment with Python 3.11
python3.11 -m venv .venv
source .venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

## Running

```bash
python main.py
```

## Project structure

```
main.py            # Entry point and home screen
view_tool.py       # PDF viewer
excerpt_tool.py    # Excerpt / region-capture tool
split_tool.py      # Split tool
pdf_to_csv_tool.py # Table extraction to CSV
library_page.py    # File library / dashboard
icons.py           # Bundled Lucide SVG icon set
colors.py          # Shared design-system colour palette
utils.py           # Shared Qt utilities
```

## License

[MIT](LICENSE)
