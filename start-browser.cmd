@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0skills\customer-express-issue-query-playwright\scripts\start_browser.ps1" %*
exit /b %errorlevel%
