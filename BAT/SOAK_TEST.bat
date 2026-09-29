@echo off
rem OFFLINE soak tests: 100-300 DOCX, vairaki merge cikli.
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"
set "PY=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY=%ROOT%\.venv\Scripts\python.exe"
echo === soak test ===
"%PY%" "%ROOT%\_diag_soak.py" %1
set "RC=%ERRORLEVEL%"
echo.
echo soak test beidzis ar kodu %RC%   (reports\SOAK_TEST_REPORT.md)
pause
endlocal & exit /b %RC%
