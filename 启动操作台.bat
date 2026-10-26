@echo off
rem ============================================================
rem  A23 Document Assistant - one-click launcher (web console)
rem  1. Double click this file.
rem  2. Browser opens at:  http://127.0.0.1:8000
rem  3. Close this window to stop the service.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
set "PY=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"

echo Starting A23 web console, browser will open in a few seconds ...
echo (keep this window open; closing it stops the service)
echo.

rem auto-open browser after ~3s; set A23_NO_BROWSER=1 to skip (for tests)
if not "%A23_NO_BROWSER%"=="1" start "" /b cmd /c "ping -n 4 127.0.0.1 >nul & start http://127.0.0.1:8000"

"%PY%" agent\api_server.py --port 8000
pause
