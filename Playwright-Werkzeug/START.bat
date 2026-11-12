@echo off
rem Playwright-Werkzeug starten (Doppelklick)
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0werkzeug.ps1"
if errorlevel 1 pause
