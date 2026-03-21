@echo off
setlocal

cd /d "%~dp0"

if /I "%~1"=="--check" goto check

where py >nul 2>nul
if not errorlevel 1 (
  py -3 ui.py
  goto end
)

where python >nul 2>nul
if not errorlevel 1 (
  python ui.py
  goto end
)

echo Python was not found. Install Python 3 and try again.
pause
exit /b 1

:check
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -m py_compile ui.py
  exit /b %errorlevel%
)

where python >nul 2>nul
if not errorlevel 1 (
  python -m py_compile ui.py
  exit /b %errorlevel%
)

echo Python was not found.
exit /b 1

:end
if errorlevel 1 (
  echo.
  echo UI exited with an error.
  pause
)
