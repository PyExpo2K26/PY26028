@echo off
title PyXcel AI Launcher
color 0A

echo ==========================================
echo        PyXcel AI-Powered Spreadsheet
echo              KiTE Development Team
echo ==========================================
echo.

:: Step 1 - Check if Ollama is installed
where ollama >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Ollama is not installed!
    echo Download it from: https://ollama.com/download
    pause
    exit /b 1
)

:: Step 2 - Check if Ollama is already running
curl -s http://localhost:11434 >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Starting Ollama in background...
    start "" /min ollama serve
    timeout /t 3 /nobreak >nul
    echo [OK] Ollama started.
) else (
    echo [OK] Ollama already running.
)

:: Step 3 - Check if llama3.1 model is pulled
echo [INFO] Checking LLaMA model...
ollama list | findstr "llama3.1" >nul 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] llama3.1 not found. Pulling now (4.7GB, needs internet once)...
    ollama pull llama3.1
    echo [OK] Model downloaded.
) else (
    echo [OK] LLaMA model ready.
)

:: Step 4 - Launch PyXcel
echo.
echo [INFO] Launching PyXcel...
echo.
python main.py

:: If Python fails
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] PyXcel failed to start.
    echo Make sure dependencies are installed:
    echo   pip install -r requirements.txt
    pause
)

