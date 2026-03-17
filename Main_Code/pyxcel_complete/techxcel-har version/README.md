# PyXcel — AI-Powered Spreadsheet System

> Built by the **KiTE Development Team**

PyXcel is a desktop application that brings AI-powered analysis, automation, and generation directly into your Excel workflow. It combines a clean PySide6 GUI with a local LLaMA model (via Ollama) to help you clean data, write formulas, build charts, generate pivot tables, and export PDFs — all without leaving the app.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Running PyXcel](#running-pyxcel)
- [Project Structure](#project-structure)
- [Module Overview](#module-overview)
- [Known Issues / Missing Modules](#known-issues--missing-modules)
- [Notes](#notes)

---

## Features

| Panel | Description |
|---|---|
|  Home | Welcome screen with Ollama status and quick-start guide |
|  Spreadsheet View | Browse and preview loaded Excel workbooks |
|  Macro Replacement | Describe a task in plain English — AI generates and runs openpyxl code |
|  Formula Generator | AI-assisted Excel formula builder with context awareness |
|  Data Cleaner | AI-generated cleaning scripts applied directly to your sheet |
|  Chat with Data | Conversational interface to query and explore your spreadsheet |
|  KPI Cards | Auto-generates key metric cards from your data |
|  Pivot Tables | Build standard or AI-guided pivot tables |
|  Chart Creator | Generate bar, line, pie, and scatter charts via matplotlib |
|  PDF Export | Export any sheet or entire workbook to a formatted PDF |
|  File Merger | Merge multiple Excel files or sheets into one workbook |

---

## Requirements

### Python Packages

```
PySide6>=6.6.0
openpyxl>=3.1.0
pandas>=2.0.0
requests>=2.31.0
matplotlib>=3.7.0
reportlab>=4.0.0
Pillow>=10.0.0
openpyxl-image-loader>=1.0.5
```

Install all at once:

```bash
pip install PySide6 pandas openpyxl matplotlib reportlab requests
```

### Python Version

Python **3.10 or higher** is recommended.

### Ollama (Local AI — Required for AI Features)

PyXcel uses [Ollama](https://ollama.com) to run a local LLaMA model. This is **not** a pip package — it must be installed separately.

1. Download and install Ollama from [https://ollama.com/download](https://ollama.com/download)
2. Pull the qwen2.5-coder:3b model (one-time download, ~1.9 GB):
   ```bash
   ollama pull qwen2.5-coder:3b
   ```
3. Start the Ollama server:
   ```bash
   ollama serve
   ```

> The app will still open without Ollama running, but all AI features will be disabled.

---

## Installation

```bash
# 1. Clone or extract the project
cd "techxcel-har version"

# 2. Install Python dependencies
pip install PySide6 pandas openpyxl matplotlib reportlab requests

# 3. Install and start Ollama (see above)
ollama pull qwen2.5-coder:3b
ollama serve
```

---

## Running PyXcel

### Windows (recommended)

Double-click `start_pyxcel.bat`. This script will:
- Check if Ollama is installed
- Start the Ollama server if not already running
- Pull the LLaMA model if not already downloaded
- Launch the app

### Any Platform (manual)

```bash
# Make sure Ollama is running first
ollama serve

# Then launch PyXcel
python main.py
```

---

## Project Structure

```
techxcel-har version/
│
├── main.py                    # Entry point — launches PySide6 app
├── start_pyxcel.bat           # Windows one-click launcher
├── requirements.txt           # Python dependencies (currently empty — see above)
│
├── core/                      # Backend logic modules
│   ├── ollama_client.py       # Ollama API communication
│   ├── workbook_inspector.py  # Excel inspection utilities
│   ├── chart_engine.py        # Chart generation (matplotlib + openpyxl)
│   ├── pivot_engine.py        # Pivot table generation
│   ├── pdf_engine.py          # PDF export (reportlab)
│   ├── code_executor.py       # Executes AI-generated openpyxl code
│  
├── gui/
│   ├── main_window.py         # Main window with sidebar navigation
│   ├── styles.qss             # Global Qt stylesheet
│   │
│   ├── panels/                # One file per sidebar panel
│   │   ├── home_panel.py
│   │   ├── spreadsheet_panel.py
│   │   ├── macro_panel.py
│   │   ├── formula_panel.py
│   │   ├── cleaner_panel.py
│   │   ├── chat_panel.py
│   │   ├── kpi_panel.py
│   │   ├── pivot_panel.py
│   │   ├── chart_panel.py
│   │   ├── pdf_panel.py
│   │   └── merger_panel.py
│   │
│   └── workers/
│       └── agent_worker.py    # QThread worker for async AI operations
│
└── uploads/                   # Auto-created at runtime
└── outputs/                   # Auto-created at runtime
```

---

## Module Overview

### `core/ollama_client.py`
Handles all communication with the local Ollama server. Exposes `is_ollama_running()`, `is_model_available()`, and `ask_llama()` which sends prompts and returns model responses.

### `core/workbook_inspector.py`
Utility functions for inspecting Excel files: `inspect_workbook()`, `get_sheet_names()`, `get_dataframe()`, `get_column_map()`, and `get_context_string()` (used to build prompts with workbook context).

### `core/chart_engine.py`
Generates bar, line, pie, and scatter charts from Excel data using matplotlib, then embeds them back into the workbook. Also supports chart preview images for the GUI.

### `core/pivot_engine.py`
Creates pivot tables from a given sheet using pandas. Supports standard pivot generation and an AI-guided mode where the instruction is sent to LLaMA to determine the best pivot configuration.

### `core/pdf_engine.py`
Exports Excel sheets to formatted PDFs using reportlab. Supports single-sheet and full-workbook export, optional summary sections, and landscape layout for wide tables.

### `core/code_executor.py`
Receives AI-generated Python (openpyxl) code, strips markdown fences, and executes it safely against the loaded workbook. Used by both the Macro and Data Cleaner panels.

### `gui/workers/agent_worker.py`
A QThread-based worker that handles all AI operations asynchronously so the UI stays responsive. Dispatches to chart, pivot, PDF, macro, cleaner, and chat workflows depending on the active panel.

---

## Known Issues / Missing Modules

The following `core/` modules are present but contain no implementation. Features that depend on them will not work until they are written:

| Module | Affects |
|---|---|
| `core/calculation.py` | Formula panel — manual calculation helpers |
| `core/data_cleaner.py` | Cleaner panel — rule-based (non-AI) cleaning |
| `core/merger.py` | Merger panel — combining workbooks/sheets |
| `core/sorter.py` | Sorting functionality across panels |
| `core/visualizer.py` | KPI cards, sparklines, visual summaries |
| `core/vlookup.py` | Formula panel — VLOOKUP generation and execution |

AI-powered fallbacks (via `code_executor.py`) still work for most of these panels, but deterministic/manual modes are non-functional.

---

## Notes

- All AI features require Ollama to be running locally on `http://localhost:11434`
- The model used is `qwen2.5-coder:3b` — this can be changed in `core/ollama_client.py`
- Uploaded files and generated outputs are saved to `uploads/` and `outputs/` folders, created automatically at launch
- The app was built and tested on Windows; the `.bat` launcher is Windows-only, but `python main.py` works cross-platform
