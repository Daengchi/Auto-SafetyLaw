@echo off
curl -L -o "%TEMP%\setup.ps1" "https://raw.githubusercontent.com/Daengchi/Auto-SafetyLaw/master/%ED%94%84%EB%A1%9C%EA%B7%B8%EB%9E%A8/setup.ps1"
PowerShell -ExecutionPolicy Bypass -File "%TEMP%\setup.ps1"
pause
