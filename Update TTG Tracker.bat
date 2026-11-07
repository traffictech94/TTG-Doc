@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "PYTHON=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "PYTHONW=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\pythonw.exe"

echo This starts the laptop Excel editor. The GitHub Pages viewer does not use this BAT file.

if not exist "%PYTHON%" (
  echo The local Python runtime used by this tracker could not be found.
  echo Please contact the person who set up the tracker.
  pause
  exit /b 1
)

"%PYTHON%" "%SCRIPT_DIR%update_tracker.py"
if errorlevel 1 (
  echo The tracker could not refresh. Check the workbook and folder names.
  pause
  exit /b 1
)

start "" "%PYTHONW%" "%SCRIPT_DIR%tracker_server.py"
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8767/"
