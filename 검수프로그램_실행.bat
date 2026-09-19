@echo off
title Urban Forest Survey Inspector
cd /d "%~dp0"

echo [1/2] Checking Python environment...
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PY_CMD=python
) else (
    where py >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set PY_CMD=py
    ) else (
        echo [ERROR] Python is not installed or not in PATH.
        echo Please install Python from https://www.python.org/
        pause
        exit /b 1
    )
)

echo [2/2] Starting Web Application Server...
%PY_CMD% web_app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application terminated unexpectedly.
    pause
)
