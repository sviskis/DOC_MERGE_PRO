@echo off
rem OFFLINE FULL testi: Word E2E, apjomi, fault injection, esosha Word sesija.
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"
set "PY=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY=%ROOT%\.venv\Scripts\python.exe"
echo === offline-test-full ===
"%PY%" -m docmerge.cli.main --offline-test-full
set "RC=%ERRORLEVEL%"
echo.
echo offline-test-full beidzis ar kodu %RC%   (reports\OFFLINE_TEST_REPORT.md)
pause
endlocal & exit /b %RC%
