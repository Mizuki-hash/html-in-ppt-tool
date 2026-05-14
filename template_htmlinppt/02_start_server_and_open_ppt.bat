@echo off
setlocal
cd /d "%~dp0"

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if %errorlevel%==0 (
  echo Please close all Microsoft PowerPoint windows first, then run this file again.
  echo This avoids Office using the old unregistered add-in state.
  pause
  exit /b 1
)

echo [1/4] Registering the local PowerPoint add-in for this Windows user...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0pd_soi_fbe_v3_local_package\setup_powerpoint_addin.ps1"
if errorlevel 1 (
  echo.
  echo Setup failed. Please right-click this file and choose Run as administrator, or ask the presenter for help.
  pause
  exit /b 1
)

echo.
echo [2/4] Checking Python 3...
where python >nul 2>nul
if %errorlevel%==0 (
  python -c "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)" >nul 2>nul
  if %errorlevel%==0 goto :python_ok
)
where python3 >nul 2>nul
if %errorlevel%==0 (
  python3 -c "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)" >nul 2>nul
  if %errorlevel%==0 goto :python_ok
)
where py >nul 2>nul
if %errorlevel%==0 (
  py -3 -c "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)" >nul 2>nul
  if %errorlevel%==0 goto :python_ok
)
echo Python 3 could not start.
echo Please install Python 3 from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
pause
exit /b 1

:python_ok
echo [3/4] Starting the local HTTPS server from this folder...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ownerPids = Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort 8765 -State Listen -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique; foreach($ownerPid in $ownerPids){ try { $proc = Get-Process -Id $ownerPid -ErrorAction Stop; if($proc.ProcessName -match 'python|py'){ Stop-Process -Id $ownerPid -Force } } catch {} }"
start "PD-SOI Local HTTPS Server" cmd /k ""%~dp0pd_soi_fbe_v3_local_package\start_https_server.bat""
timeout /t 3 /nobreak >nul

echo [4/4] Opening with Microsoft PowerPoint...
set "PPT=%~dp0PD_SOI_FBE_V3_Direct_In_PowerPoint.pptx"
set "POWERPNT=C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE"
if exist "%POWERPNT%" (
  start "" "%POWERPNT%" "%PPT%"
  exit /b
)

set "POWERPNT=C:\Program Files (x86)\Microsoft Office\Root\Office16\POWERPNT.EXE"
if exist "%POWERPNT%" (
  start "" "%POWERPNT%" "%PPT%"
  exit /b
)

where POWERPNT.EXE >nul 2>nul
if %errorlevel%==0 (
  start "" POWERPNT.EXE "%PPT%"
  exit /b
)

echo Microsoft PowerPoint was not found.
echo WPS cannot run this embedded Office web add-in interactively.
pause
