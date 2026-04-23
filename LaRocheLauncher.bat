@echo off

:: Vérifie si le script tourne en admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Demande des droits administrateur...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Si on est admin, on lance le script PowerShell
PowerShell -ExecutionPolicy Bypass -File "%~dp0LaRoche.ps1"
pause