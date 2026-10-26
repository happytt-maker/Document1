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
set "PY=%USERPROFILE%\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

"%PY%" "!SCRIPT!" %*
echo.
pause
