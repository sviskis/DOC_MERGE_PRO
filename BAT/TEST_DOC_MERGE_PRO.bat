@echo off
rem Unit/offline testi ar pytest.
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"
set "PY=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY=%ROOT%\.venv\Scripts\python.exe"
echo === pytest -q ===
"%PY%" -m pytest -q
set "RC=%ERRORLEVEL%"
echo.
echo pytest beidzis ar kodu %RC%
pause
endlocal & exit /b %RC%
