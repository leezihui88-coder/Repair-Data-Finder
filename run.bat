@echo off
title MOXA Schematic Link Finder v1.4

echo [1/2] Preparing environment with uv...
echo [2/2] Launching main program (v1.4)...
echo.

uv run python "%~dp0moxa_schematic_finder_v1.4.py"

if errorlevel 1 (
    echo.
    echo [ERROR] Program exited abnormally. Please screenshot this window.
    pause
)
