@echo off
setlocal EnableExtensions DisableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0" || (
    echo Failed to enter the toolkit folder.
    pause
    exit /b 1
)

REM Force Python/conda subprocesses to use UTF-8 to avoid
REM UnicodeDecodeError with Chinese characters in paths (e.g. 自动化).
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ==========================================
echo Account Audit Toolkit - Conda Installation
echo ==========================================

set "APP_DIR=%~dp0"
set "ENV_DIR=%APP_DIR%.conda_env"
set "PYTHON_EXE=%ENV_DIR%\python.exe"
set "REQ_FILE=%APP_DIR%requirements.txt"

if not exist "%REQ_FILE%" (
    echo requirements.txt was not found in:
    echo %APP_DIR%
    pause
    exit /b 1
)

call :find_conda
if not defined CONDA_CMD (
    call :install_miniconda
    if errorlevel 1 goto conda_failed
    call :find_conda
)

if not defined CONDA_CMD (
    echo Conda was not found after installation.
    pause
    exit /b 1
)

echo Conda command: %CONDA_CMD%
call "%CONDA_CMD%" --version
if errorlevel 1 goto conda_failed

if not exist "%PYTHON_EXE%" goto create_env

"%PYTHON_EXE%" -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 12) else 1)"
if errorlevel 1 (
    echo The existing .conda_env is not Python 3.12.
    echo Delete the .conda_env folder and run this file again.
    pause
    exit /b 1
)
goto env_ready

:create_env
echo.
echo Creating local Conda environment with Python 3.12...
call "%CONDA_CMD%" create --prefix "%ENV_DIR%" python=3.12 pip -y
if errorlevel 1 goto conda_failed

:env_ready

if not exist "%PYTHON_EXE%" (
    echo Failed to create the Conda environment.
    pause
    exit /b 1
)

echo.
echo Upgrading pip...
call "%CONDA_CMD%" run --prefix "%ENV_DIR%" --no-capture-output python -m pip install --upgrade pip
if errorlevel 1 goto install_failed

echo.
echo Installing project requirements...
call "%CONDA_CMD%" run --prefix "%ENV_DIR%" --no-capture-output python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn -r "%REQ_FILE%"
if errorlevel 1 (
    echo.
    echo Tsinghua mirror installation failed. Retrying with the default PyPI index...
    call "%CONDA_CMD%" run --prefix "%ENV_DIR%" --no-capture-output python -m pip install -r "%REQ_FILE%"
    if errorlevel 1 goto install_failed
)

echo.
echo Verifying the installation...
call "%CONDA_CMD%" run --prefix "%ENV_DIR%" --no-capture-output python -c "import flask, pandas, openpyxl, xlrd, numpy, PIL, rapidocr_onnxruntime; print('All required packages are available.')"
if errorlevel 1 goto install_failed

echo.
echo Installation completed. Run the start BAT file to open the tool.
pause
exit /b 0

:find_conda
set "CONDA_CMD="
where conda >nul 2>nul
if not errorlevel 1 (
    set "CONDA_CMD=conda"
    exit /b 0
)

if exist "%USERPROFILE%\miniconda3\condabin\conda.bat" set "CONDA_CMD=%USERPROFILE%\miniconda3\condabin\conda.bat"
if not defined CONDA_CMD if exist "%USERPROFILE%\anaconda3\condabin\conda.bat" set "CONDA_CMD=%USERPROFILE%\anaconda3\condabin\conda.bat"
if not defined CONDA_CMD if exist "%LOCALAPPDATA%\miniconda3\condabin\conda.bat" set "CONDA_CMD=%LOCALAPPDATA%\miniconda3\condabin\conda.bat"
if not defined CONDA_CMD if exist "%LOCALAPPDATA%\anaconda3\condabin\conda.bat" set "CONDA_CMD=%LOCALAPPDATA%\anaconda3\condabin\conda.bat"
if not defined CONDA_CMD if exist "%ProgramData%\miniconda3\condabin\conda.bat" set "CONDA_CMD=%ProgramData%\miniconda3\condabin\conda.bat"
if not defined CONDA_CMD if exist "%ProgramData%\anaconda3\condabin\conda.bat" set "CONDA_CMD=%ProgramData%\anaconda3\condabin\conda.bat"

if not defined CONDA_CMD if exist "%USERPROFILE%\miniconda3\Scripts\conda.exe" set "CONDA_CMD=%USERPROFILE%\miniconda3\Scripts\conda.exe"
if not defined CONDA_CMD if exist "%USERPROFILE%\anaconda3\Scripts\conda.exe" set "CONDA_CMD=%USERPROFILE%\anaconda3\Scripts\conda.exe"
if not defined CONDA_CMD if exist "%LOCALAPPDATA%\miniconda3\Scripts\conda.exe" set "CONDA_CMD=%LOCALAPPDATA%\miniconda3\Scripts\conda.exe"
if not defined CONDA_CMD if exist "%LOCALAPPDATA%\anaconda3\Scripts\conda.exe" set "CONDA_CMD=%LOCALAPPDATA%\anaconda3\Scripts\conda.exe"
if not defined CONDA_CMD if exist "%ProgramData%\miniconda3\Scripts\conda.exe" set "CONDA_CMD=%ProgramData%\miniconda3\Scripts\conda.exe"
if not defined CONDA_CMD if exist "%ProgramData%\anaconda3\Scripts\conda.exe" set "CONDA_CMD=%ProgramData%\anaconda3\Scripts\conda.exe"
exit /b 0

:install_miniconda
echo Conda was not found. Attempting automatic download and installation of Miniconda...
set "MINICONDA_URL=https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe"
set "MINICONDA_INSTALLER=%TEMP%\Miniconda3_installer_%RANDOM%.exe"
set "MINICONDA_DIR=%USERPROFILE%\miniconda3"

echo Downloading Miniconda from %MINICONDA_URL%...
powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%MINICONDA_URL%' -OutFile '%MINICONDA_INSTALLER%'"
if errorlevel 1 (
    echo Failed to download Miniconda. Please check your network connection.
    del "%MINICONDA_INSTALLER%" 2>nul
    exit /b 1
)

echo Installing Miniconda silently to %MINICONDA_DIR% ...
"%MINICONDA_INSTALLER%" /InstallationType=JustMe /RegisterPython=0 /S /D=%MINICONDA_DIR%
set "INSTALL_CODE=%ERRORLEVEL%"
del "%MINICONDA_INSTALLER%" 2>nul
if not "%INSTALL_CODE%"=="0" (
    echo Failed to install Miniconda. Installer exit code: %INSTALL_CODE%
    exit /b 1
)

if exist "%MINICONDA_DIR%\condabin\conda.bat" (
    echo Miniconda installed successfully. Continuing with environment setup...
    exit /b 0
)
if exist "%MINICONDA_DIR%\Scripts\conda.exe" (
    echo Miniconda installed successfully. Continuing with environment setup...
    exit /b 0
)

echo Miniconda installation completed but conda was not found at the expected location.
exit /b 1

:conda_failed
echo.
echo Conda environment creation failed.
echo Check the Conda installation, network, proxy, and available disk space.
pause
exit /b 1

:install_failed
echo.
echo Requirement installation failed.
echo Check the network, proxy, Python package mirror, or contact IT for offline package installation.
pause
exit /b 1
