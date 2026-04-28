@echo off
setlocal
cd /d "%~dp0"

if not exist "%~dp0assets\RutherfordLauncher.ps1" (
  echo Missing assets\RutherfordLauncher.ps1 next to the launcher.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0assets\RutherfordLauncher.ps1"
exit /b %ERRORLEVEL%
