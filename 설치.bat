@echo off
curl -L -o "%TEMP%\setup.ps1" "https://raw.githubusercontent.com/Daengchi/Auto-SafetyLaw/master/setup.ps1"
PowerShell -ExecutionPolicy Bypass -File "%TEMP%\setup.ps1"
