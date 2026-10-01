@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0check.ps1"
set "task_exit=%errorlevel%"
if not "%~1"=="--no-pause" pause
exit /b %task_exit%
