@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM  Fully automated build script for JScanner.py
REM  -> creates a folder-based Windows app (dist\JScanner)
REM
REM  What this does automatically, with no manual steps:
REM   1. Elevates itself to Administrator (needed to install
REM      Python system-wide).
REM   2. Checks if Python is installed; installs it via winget,
REM      or falls back to a silent install from python.org.
REM   3. Refreshes PATH in the current window (no need to close
REM      and reopen anything).
REM   4. Creates an isolated virtual environment (.venv).
REM   5. Installs all required packages inside that venv.
REM   6. Runs PyInstaller to build the app.
REM   7. Creates a shortcut to the built .exe on the Desktop and
REM      in the Start Menu.
REM
REM  Just double-click this file. Run it from the folder that
REM  contains JScanner.py, requirements.txt and icon.ico.
REM ============================================================

cd /d "%~dp0"

REM --- Self-elevate to Administrator if not already elevated ---
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrator privileges...
    powershell -Command "Start-Process -FilePath '%~f0' -WorkingDirectory '%~dp0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo  Step 1: Checking for Python...
echo ============================================================

set "PYTHON_CMD="

py -3 --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :python_found
)

python --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo Python was not found on this system. Installing it now...
echo.

set "PY_INSTALLED=0"

where winget >nul 2>&1
if %errorlevel%==0 (
    winget install -e --id Python.Python.3.12 --scope machine --accept-source-agreements --accept-package-agreements --silent
    if !errorlevel!==0 set "PY_INSTALLED=1"
)

if "!PY_INSTALLED!"=="0" (
    echo winget install failed or unavailable. Downloading the official
    echo installer from python.org instead...
    set "PY_INSTALLER=%TEMP%\python-installer.exe"
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.10/python-3.12.10-amd64.exe' -OutFile '%PY_INSTALLER%'"
    if !errorlevel! neq 0 (
        echo ERROR: Could not download the Python installer.
        echo Please install Python manually from https://www.python.org/downloads/
        echo IMPORTANT: check "Add python.exe to PATH" during setup, then run this script again.
        pause
        exit /b 1
    )

    "!PY_INSTALLER!" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    if !errorlevel! neq 0 (
        echo ERROR: Automatic installation of Python failed.
        echo Please install Python manually from https://www.python.org/downloads/
        echo IMPORTANT: check "Add python.exe to PATH" during setup, then run this script again.
        pause
        exit /b 1
    )
    del "!PY_INSTALLER!" >nul 2>&1
)

echo Python installed. Refreshing PATH for this session...

REM --- Reload PATH from the registry (Machine + User) without
REM     needing to close and reopen the window ---
for /f "usebackq skip=2 tokens=2,*" %%A in (`reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul`) do set "SysPath=%%B"
for /f "usebackq skip=2 tokens=2,*" %%A in (`reg query "HKCU\Environment" /v Path 2^>nul`) do set "UserPath=%%B"
set "PATH=%SysPath%;%UserPath%"

REM --- Retry finding Python now that PATH is refreshed ---
py -3 --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=py -3"
    goto :python_found
)
python --version >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto :python_found
)

REM --- Last resort: check the well-known default install paths ---
if exist "C:\Program Files\Python312\python.exe" (
    set "PYTHON_CMD=C:\Program Files\Python312\python.exe"
    goto :python_found
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto :python_found
)

echo ERROR: Python was installed but could not be detected automatically.
echo Please close this window, open a new Command Prompt, and run this script again.
pause
exit /b 1

:python_found
echo Found Python using command: %PYTHON_CMD%
%PYTHON_CMD% --version

echo.
echo ============================================================
echo  Step 2: Creating virtual environment (.venv)...
echo ============================================================
if not exist ".venv" (
    %PYTHON_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create the virtual environment.
        pause
        exit /b 1
    )
) else (
    echo Virtual environment already exists, skipping creation.
)

set "VENV_PY=.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
    echo ERROR: Virtual environment python.exe not found at %VENV_PY%
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Step 3: Installing/updating required packages...
echo ============================================================
"%VENV_PY%" -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo ERROR: Failed to upgrade pip.
    pause
    exit /b 1
)

"%VENV_PY%" -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Failed to install required packages from requirements.txt
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Step 4: Building folder-based app with PyInstaller...
echo ============================================================
"%VENV_PY%" -m PyInstaller --onedir --windowed --noconfirm --clean ^
    --name "JScanner" ^
    --icon "icon.ico" ^
    --add-data "icon.ico;." ^
    --collect-submodules win32com ^
    --hidden-import win32timezone ^
    JScanner.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: PyInstaller build failed. See the messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Step 5: Creating shortcuts (Desktop + Start Menu)...
echo ============================================================
set "EXE_PATH=%~dp0dist\JScanner\JScanner.exe"
set "ICON_PATH=%~dp0dist\JScanner\icon.ico"
if not exist "%ICON_PATH%" set "ICON_PATH=%EXE_PATH%"
set "WORK_DIR=%~dp0dist\JScanner"
set "SHORTCUT_PS1=%TEMP%\jscanner_make_shortcut.ps1"

del "%SHORTCUT_PS1%" >nul 2>&1
echo $ws = New-Object -ComObject WScript.Shell> "%SHORTCUT_PS1%"
echo $exe = "%EXE_PATH%">> "%SHORTCUT_PS1%"
echo $icon = "%ICON_PATH%">> "%SHORTCUT_PS1%"
echo $workDir = "%WORK_DIR%">> "%SHORTCUT_PS1%"
echo $desktopPath = [Environment]::GetFolderPath^('Desktop'^) + '\JScanner.lnk'>> "%SHORTCUT_PS1%"
echo $startMenuPath = [Environment]::GetFolderPath^('StartMenu'^) + '\Programs\JScanner.lnk'>> "%SHORTCUT_PS1%"
echo $s1 = $ws.CreateShortcut^($desktopPath^)>> "%SHORTCUT_PS1%"
echo $s1.TargetPath = $exe>> "%SHORTCUT_PS1%"
echo $s1.WorkingDirectory = $workDir>> "%SHORTCUT_PS1%"
echo $s1.IconLocation = $icon>> "%SHORTCUT_PS1%"
echo $s1.Save^(^)>> "%SHORTCUT_PS1%"
echo $s2 = $ws.CreateShortcut^($startMenuPath^)>> "%SHORTCUT_PS1%"
echo $s2.TargetPath = $exe>> "%SHORTCUT_PS1%"
echo $s2.WorkingDirectory = $workDir>> "%SHORTCUT_PS1%"
echo $s2.IconLocation = $icon>> "%SHORTCUT_PS1%"
echo $s2.Save^(^)>> "%SHORTCUT_PS1%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%SHORTCUT_PS1%"
set "SHORTCUT_RESULT=%errorlevel%"
del "%SHORTCUT_PS1%" >nul 2>&1

if %SHORTCUT_RESULT% neq 0 (
    echo WARNING: Could not create shortcuts automatically. You can
    echo still run the app directly from dist\JScanner\JScanner.exe
) else (
    echo Shortcuts created on the Desktop and in the Start Menu.
)

echo.
echo ============================================================
echo  Done!
echo ============================================================
echo The app is in the "dist\JScanner" folder.
echo JScanner.exe there is small and starts instantly, but it
echo needs the other files next to it - copy/zip the WHOLE folder,
echo not just the .exe.
echo A shortcut has also been added to your Desktop and Start Menu.
pause
