@echo off
setlocal
cd /d "%~dp0"

echo === Project Planer LXC Publish ===
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0PUBLISH.ps1" %*
set ERR=%ERRORLEVEL%

echo.
if not "%ERR%"=="0" (
  echo Publish failed with exit code %ERR%.
  pause
  exit /b %ERR%
)

echo Publish completed successfully.
pause
