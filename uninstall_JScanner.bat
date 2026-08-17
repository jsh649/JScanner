@echo off
setlocal enabledelayedexpansion
REM ============================================================
REM  Uninstaller for JScanner
REM
REM  This removes everything that build_exe_onedir.bat created:
REM   - Desktop shortcut
REM   - Start Menu shortcut
REM   - The built app folder ("dist\JScanner")
REM   - PyInstaller's intermediate "build" folder and .spec file
REM   - The virtual environment (".venv"), which removes ALL
REM     Python packages that were installed for this project
REM     (PyQt5, opencv-python, numpy, Pillow, pywin32, pyinstaller)
REM
REM  Python itself is NOT removed/uninstalled, since it may be
REM  used by other programs on this computer.
REM
REM  This does NOT touch JScanner.py, requirements.txt, or
REM  icon.ico - your source files are kept.
REM ============================================================

cd /d "%~dp0"

echo ============================================================
echo  JScanner Uninstaller
echo ============================================================
echo This will delete:
echo   - Desktop shortcut
echo   - Start Menu shortcut
echo   - dist\JScanner  (the built app)
echo   - build          (PyInstaller temp files)
echo   - JScanner.spec  (PyInstaller config file)
echo   - .venv          (isolated Python environment and all
echo                      packages installed for this project)
echo.
echo Python itself will NOT be removed.
echo Your source files (JScanner.py, requirements.txt, icon.ico)
echo will NOT be removed.
echo.
set /p CONFIRM="Are you sure you want to continue? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo Cancelled. Nothing was removed.
    pause
    exit /b 0
)

echo.
echo ============================================================
echo  Removing shortcuts...
echo ============================================================
set "SHORTCUT_PS1=%TEMP%\jscanner_remove_shortcut.ps1"
del "%SHORTCUT_PS1%" >nul 2>&1
echo $desktopPath = [Environment]::GetFolderPath^('Desktop'^) + '\JScanner.lnk'> "%SHORTCUT_PS1%"
echo $startMenuPath = [Environment]::GetFolderPath^('StartMenu'^) + '\Programs\JScanner.lnk'>> "%SHORTCUT_PS1%"
echo if ^(Test-Path $desktopPath^) { Remove-Item $desktopPath -Force }>> "%SHORTCUT_PS1%"
echo if ^(Test-Path $startMenuPath^) { Remove-Item $startMenuPath -Force }>> "%SHORTCUT_PS1%"

powershell -NoProfile -ExecutionPolicy Bypass -File "%SHORTCUT_PS1%"
del "%SHORTCUT_PS1%" >nul 2>&1
echo Shortcuts removed (if they existed).

echo.
echo ============================================================
echo  Removing built app (dist\JScanner)...
echo ============================================================
if exist "dist\JScanner" (
    rmdir /s /q "dist\JScanner"
    echo Removed dist\JScanner
) else (
    echo dist\JScanner not found, skipping.
)
REM Remove the "dist" folder itself only if it is now empty
if exist "dist" (
    rmdir "dist" 2>nul
)

echo.
echo ============================================================
echo  Removing PyInstaller temp files (build, .spec)...
echo ============================================================
if exist "build" (
    rmdir /s /q "build"
    echo Removed build\
) else (
    echo build\ not found, skipping.
)
if exist "JScanner.spec" (
    del /q "JScanner.spec"
    echo Removed JScanner.spec
) else (
    echo JScanner.spec not found, skipping.
)

echo.
echo ============================================================
echo  Removing virtual environment (.venv) and its packages...
echo ============================================================
if exist ".venv" (
    rmdir /s /q ".venv"
    echo Removed .venv and all packages installed inside it
    echo ^(PyQt5, opencv-python, numpy, Pillow, pywin32, pyinstaller^).
) else (
    echo .venv not found, skipping.
)

echo.
echo ============================================================
echo  Done!
echo ============================================================
echo JScanner has been uninstalled.
echo Python itself was left untouched, since it may be shared
echo with other programs on this computer.
echo Your source files (JScanner.py, requirements.txt, icon.ico)
echo were kept in this folder.
pause
