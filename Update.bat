@echo off
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0update.ps1" -Env test
pause

