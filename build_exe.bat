@echo off
setlocal

echo ============================================================
echo  Building a portable EXE (one-time setup)
echo  This step needs Python. The EXE it produces will NOT.
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found on this computer.
    echo.
    echo This one-time build step needs Python installed, just to
    echo run the packaging tool. Install it from:
    echo   https://www.python.org/downloads/
    echo During install, check the box "Add python.exe to PATH".
    echo Then run this file again. You only need to do this once,
    echo on one computer.
    echo.
    pause
    exit /b 1
)

echo [1/5] Creating a throwaway virtual environment ".buildenv" ...
python -m venv .buildenv
if errorlevel 1 goto :fail

echo [2/5] Installing flask, openpyxl and pyinstaller ...
.buildenv\Scripts\python.exe -m pip install --upgrade pip -q
if errorlevel 1 goto :fail
.buildenv\Scripts\python.exe -m pip install -r requirements.txt pyinstaller -q
if errorlevel 1 goto :fail

echo [3/5] Building the EXE, this can take a minute or two ...
.buildenv\Scripts\python.exe -m PyInstaller --onefile --console --name AttendanceViewer --add-data "%~dp0templates;templates" --add-data "%~dp0static;static" --distpath . --workpath .buildwork --specpath .buildwork app.py
if errorlevel 1 goto :fail

echo [4/5] Cleaning up build files ...
rmdir /s /q .buildwork >nul 2>nul
rmdir /s /q .buildenv >nul 2>nul

echo [5/5] Done!
echo.
echo ============================================================
echo  AttendanceViewer.exe is ready, right here in this folder.
echo.
echo  It is fully portable: copy it, together with a "Result"
echo  folder placed next to it, to any Windows computer and just
echo  double-click it. No Python required on that computer.
echo ============================================================
echo.
pause
exit /b 0

:fail
echo.
echo [ERROR] Something went wrong during the build. See the messages above.
rmdir /s /q .buildwork >nul 2>nul
rmdir /s /q .buildenv >nul 2>nul
pause
exit /b 1
