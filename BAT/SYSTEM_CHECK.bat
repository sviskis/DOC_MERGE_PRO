@echo off
rem Word COM un vides parbaude (--system-check).
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"
set "PY=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY=%ROOT%\.venv\Scripts\python.exe"
echo === system-check ===
"%PY%" -m docmerge.cli.main --system-check
set "RC=%ERRORLEVEL%"
echo.
echo system-check beidzis ar kodu %RC%
pause
endlocal & exit /b %RC%
