@echo off
title PT NIFCO INDONESIA - MILKRUN LAUNCHER
color 0B
cls

echo ============================================================
echo        PT NIFCO INDONESIA - MILKRUN DASHBOARD LAUNCHER
echo                 Windows PowerShell / CMD
echo ============================================================
echo.

cd /d "%~dp0"

:: 1. Deteksi Python
echo [*] Mendeteksi Python di komputer...
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    where py >nul 2>&1
    if %ERRORLEVEL% NEQ 0 (
        if exist "%LOCALAPPDATA%\Python\bin\python.exe" (
            set PY_CMD="%LOCALAPPDATA%\Python\bin\python.exe"
            goto python_found
        )
        if exist "Instal Python\python-3.15.0-amd64.exe" (
            color 0E
            echo [!] Python belum terpasang di komputer ini.
            echo [*] Ditemukan file installer di folder "Instal Python".
            echo [*] Menjalankan installer Python...
            start /wait "" "Instal Python\python-3.15.0-amd64.exe"
            echo [*] Instalasi selesai. Silakan restart launcher.
            pause
            exit /b 0
        )
        color 0C
        echo [ERROR] Python tidak terdeteksi di komputer ini!
        echo.
        echo Silakan download dan install Python terlebih dahulu:
        echo https://www.python.org/downloads/
        echo (Pastikan centang "Add Python to PATH" saat menginstall)
        echo.
        pause
        exit /b 1
    ) else (
        set PY_CMD=py
    )
) else (
    set PY_CMD=python
)

:python_found
for /f "tokens=*" %%i in ('%PY_CMD% --version 2^>^&1') do set PY_VER=%%i
echo [OK] Menggunakan: %PY_VER% (%PY_CMD%)

:: 2. Jalankan launcher.py
%PY_CMD% "%~dp0launcher.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Server berhenti dengan status error (%ERRORLEVEL%).
    pause
)
