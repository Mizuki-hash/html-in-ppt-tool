@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

if "%~1"=="" (
  echo Drag generated htmlinppt folders or generated .pptx files onto this BAT.
  echo.
  echo You can also paste paths one by one. Leave an empty line to start merging.
  echo.
  set "ARGS="
  :read_loop
  set /p "ONE=Path: "
  if "%ONE%"=="" goto :run_merge
  set ARGS=!ARGS! "%ONE%"
  goto :read_loop
) else (
  set ARGS=%*
)

:run_merge
where python >nul 2>nul
if %errorlevel%==0 (
  python "%~dp0merge_html_ppt_packages.py" !ARGS!
  pause
  exit /b %errorlevel%
)

where python3 >nul 2>nul
if %errorlevel%==0 (
  python3 "%~dp0merge_html_ppt_packages.py" !ARGS!
  pause
  exit /b %errorlevel%
)

where py >nul 2>nul
if %errorlevel%==0 (
  py -3 "%~dp0merge_html_ppt_packages.py" !ARGS!
  pause
  exit /b %errorlevel%
)

echo Python 3 could not start.
pause
exit /b 1
