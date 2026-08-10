@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\refresh_followup_only.ps1"
exit /b %ERRORLEVEL%
