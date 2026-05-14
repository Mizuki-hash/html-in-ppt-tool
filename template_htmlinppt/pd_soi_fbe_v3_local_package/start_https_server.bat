@echo off
setlocal
cd /d "%~dp0"
title PD-SOI FBE V3 HTTPS Local Server

set "SCRIPT=%~dp0server_https.py"

echo Starting local HTTPS server...
echo Folder: %~dp0
echo.

where python >nul 2>nul
if %errorlevel%==0 (
  echo Trying: python
  python "%SCRIPT%"
  if not errorlevel 1 goto :end
  echo python failed, trying the next option...
  echo.
)

where python3 >nul 2>nul
if %errorlevel%==0 (
  echo Trying: python3
  python3 "%SCRIPT%"
  if not errorlevel 1 goto :end
  echo python3 failed, trying the next option...
  echo.
)

where py >nul 2>nul
if %errorlevel%==0 (
  echo Trying: py -3
  py -3 "%SCRIPT%"
  if not errorlevel 1 goto :end
  echo py -3 failed.
  echo.
)

echo Python 3 could not start.
echo If Python is installed, check that python.exe works in Command Prompt.
echo Otherwise install Python 3 from https://www.python.org/downloads/ and tick "Add python.exe to PATH".

:end
pause
