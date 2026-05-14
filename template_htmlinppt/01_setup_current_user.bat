@echo off
setlocal
cd /d "%~dp0"

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if %errorlevel%==0 (
  echo Please close all Microsoft PowerPoint windows first, then run this file again.
  pause
  exit /b 1
)

echo This registers the local PowerPoint web add-in for the current Windows user.
echo You normally only need to run 02_start_server_and_open_ppt.bat.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0pd_soi_fbe_v3_local_package\setup_powerpoint_addin.ps1"
if errorlevel 1 (
  echo.
  echo Setup failed. Try right-clicking this file and choosing Run as administrator.
  pause
  exit /b 1
)

echo.
echo Setup done. Close PowerPoint, then run 02_start_server_and_open_ppt.bat.
pause
