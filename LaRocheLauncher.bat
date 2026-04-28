@echo off
setlocal
cd /d "%~dp0"

if exist "%~dp0RutherfordAssistant.exe" (
  start "" "%~dp0RutherfordAssistant.exe"
  exit /b 0
)

if not exist "%~dp0assets\RutherfordLauncher.ps1" (
  echo Missing assets\RutherfordLauncher.ps1 next to the launcher.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath 'powershell.exe' -ArgumentList '-NoProfile -ExecutionPolicy Bypass -STA -File ""%~dp0assets\RutherfordLauncher.ps1""' -Verb RunAs"
exit /b 0
