PyXcel Complete (TechXcel-HAR Version)

AI-powered desktop application for spreadsheet analysis, automation, and reporting.

Key features:
- Spreadsheet inspection and cleaning
- Formula and macro assistance
- Chat with data (Ollama/LLaMA integration)
- KPI cards, chart generation, and pivot tools
- File merge and PDF export
- Premium UI refresh with dual themes

New UI enhancements:
- Premium visual polish across panels and sidebar
- Theme toggle button in the left sidebar
- Full light and dark theme support
- Saved theme preference (restored on next app launch)

Project structure:
- main.py: application entry point
- core/: data and analysis engines
- gui/: UI components, styles, worker threads
- gui/styles.qss: premium dark theme
- gui/styles_light.qss: premium light theme
- requirements.txt: Python dependencies
- start_pyxcel.bat: Windows launcher helper

Prerequisites:
- Windows 10 or later
- Python 3.10+ recommended

Setup:
1. Open a terminal in this folder:
   k:\Py-expo25\Final_final\PY26028\Main_Code\pyxcel_complete\techxcel-har version
2. Optional: create and activate a virtual environment.
3. Install dependencies:
   pip install -r requirements.txt

Run:
- Option 1:
  python main.py
- Option 2:
  start_pyxcel.bat

Using theme toggle:
1. Launch the app.
2. In the sidebar, click the theme button below "Load Excel File".
3. Click again to switch between light and dark modes.
4. The selected theme is remembered automatically.

Troubleshooting:
- If imports fail, run `pip install -r requirements.txt` again in the same environment used to run the app.
- If AI features are unavailable, start Ollama first (`ollama serve`) and ensure model `llama3.1` is pulled.
- If GUI fails to open, verify PySide6 is installed in your active Python environment.

Notes:
- Keep `core` and `gui` folders in place relative to `main.py`.
- Use the same Python interpreter for install and run steps.
