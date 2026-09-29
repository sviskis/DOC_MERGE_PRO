@echo off
rem OFFLINE QUICK testi (bez tikla, prasa tikai lokalo Python).
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"
set "PY=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY=%ROOT%\.venv\Scripts\python.exe"
echo === offline-test (QUICK) ===
"%PY%" -m docmerge.cli.main --offline-test
set "RC=%ERRORLEVEL%"
echo.
echo offline-test beidzis ar kodu %RC%   (reports\OFFLINE_QUICK_TEST_REPORT.md)
pause
endlocal & exit /b %RC%
