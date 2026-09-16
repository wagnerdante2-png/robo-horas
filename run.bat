@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  call setup.bat
  if errorlevel 1 exit /b 1
)

".venv\Scripts\python.exe" main.py %*
set EXITCODE=%ERRORLEVEL%

echo.
if not "%EXITCODE%"=="0" (
  echo O robo terminou com erro. Consulte output\logs.
  pause
)

exit /b %EXITCODE%
