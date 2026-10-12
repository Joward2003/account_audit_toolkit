@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0" || (
    echo Failed to enter the toolkit folder.
    pause
    exit /b 1
)

REM Force Python to use UTF-8 to avoid
REM UnicodeDecodeError with Chinese characters in paths.
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

set "APP_DIR=%~dp0"
set "PYTHON_EXE=%APP_DIR%.conda_env\python.exe"
set "WEB_APP=%APP_DIR%web_app.py"
set "APP_URL=http://127.0.0.1:5000"

if not exist "%PYTHON_EXE%" (
    echo Please run the install BAT file first.
    pause
    exit /b 1
)

if not exist "%WEB_APP%" (
    echo web_app.py was not found in:
    echo %APP_DIR%
    pause
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "if (Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort 5050 -State Listen -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }" >nul 2>nul
if errorlevel 1 (
    echo Port 5050 is already in use.
    echo Close the other program using this port, or stop the previous toolkit window, then run this file again.
    pause
    exit /b 1
)

echo Starting Account Audit Toolkit...
echo Keep this window open while using the tool.
echo Press Ctrl+C to stop the tool.
echo URL: %APP_URL%

start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "Start-Sleep -Seconds 2; Start-Process '%APP_URL%'"
"%PYTHON_EXE%" "%WEB_APP%"

echo.
echo The toolkit has stopped.
pause
