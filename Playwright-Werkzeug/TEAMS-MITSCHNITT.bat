@echo off
rem Direkt den Teams-Mitschnitt starten (Doppelklick, z. B. kurz vor dem Termin)
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0werkzeug.ps1" -Teams
if errorlevel 1 pause
