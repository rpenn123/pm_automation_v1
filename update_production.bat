@echo off
powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0update_prod.ps1" -Env prod
pause

