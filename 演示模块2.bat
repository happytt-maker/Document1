@echo off
rem ============================================================
rem  A23 Module 2 - full pipeline demo (double click to run)
rem  1. Double click this file.
rem  2. It runs: pool -> extract x3 -> store -> quality check
rem     (about 25 seconds; --fast gives a 2-second quick view)
rem ============================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"
set "SCRIPT=agent\ÑÝÊ¾-Ä£¿é2.py"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

rem --- interpreter: prefer the original dev env, fall back to python on PATH ---
set "PY=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if exist "%PY%" goto have_py
set "PY=python"
where python >nul 2>&1 || goto no_python

:have_py
"%PY%" "!SCRIPT!" %*
echo.
pause
exit /b

:no_python
echo [ERROR] Python not found on this machine.
echo Install Python 3.10+ from python.org, tick "Add python.exe to PATH", then run again.
pause
exit /b 1
