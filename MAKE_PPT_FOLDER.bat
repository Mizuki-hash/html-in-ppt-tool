@echo off
setlocal
cd /d "%~dp0"

set "HTML=%~1"
if "%HTML%"=="" (
  echo Drag one .html file onto this BAT, or paste the html file path below.
  echo.
  set /p "HTML=HTML file path: "
)

if not exist "%HTML%" (
  echo.
  echo File not found:
  echo %HTML%
  pause
  exit /b 1
)

where python >nul 2>nul
if %errorlevel%==0 (
  python "%~dp0html_to_ppt_packager.py" "%HTML%"
  pause
  exit /b %errorlevel%
)

where python3 >nul 2>nul
if %errorlevel%==0 (
  python3 "%~dp0html_to_ppt_packager.py" "%HTML%"
  pause
  exit /b %errorlevel%
)

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%~dp0html_to_ppt_packager.py" "%HTML%"
  pause
  exit /b %errorlevel%
)

echo Python 3 could not start.
echo Install Python 3 from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
pause
exit /b 1
