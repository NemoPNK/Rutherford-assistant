@echo off
setlocal
cd /d "%~dp0"

:: If relaunched with argument (after elevation)
if /I "%~1"=="setup" goto run_setup
if /I "%~1"=="network" goto run_network

:menu
echo.
echo ======================================
echo      Rutherford assistant launcher
echo ======================================
echo.
echo Available commands:
echo   s   = setup / run LaRoche.ps1
echo   n = network / run Network.ps1 NE PAS UTILISER CE N'EST PAS FINI 
echo   exit    = quit

echo.
set /p userChoice=Enter command (s/n/exit): 

if /I "%userChoice%"=="s" goto elevate_setup
if /I "%userChoice%"=="n" goto elevate_network
if /I "%userChoice%"=="exit" goto end

echo.
echo Invalid command.
pause
goto menu

:elevate_setup
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList 'setup' -Verb RunAs"
exit /b

:elevate_network
powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList 'network' -Verb RunAs"
exit /b

:run_setup
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0LaRoche.ps1"
pause
goto end

:run_network
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Network.ps1"
pause
goto end

:end
exit /b