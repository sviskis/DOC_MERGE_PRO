@echo off
rem DOC_MERGE_PRO galvenais GUI launcher.
rem Strada neatkarigi no tas, no kuras mapes to palaiz (%TEMP%, C:\, u.t.t.).
setlocal
set "ROOT=%~dp0.."
cd /d "%ROOT%"
set "PY=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY=%ROOT%\.venv\Scripts\python.exe"
echo === DOC_MERGE_PRO ===
echo Projekta mape: %ROOT%
echo Python: %PY%
"%PY%" "%ROOT%\DOC_MERGE_PRO.py"
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
  echo.
  echo KLIUDA: DOC_MERGE_PRO beidzas ar kodu %RC%
  echo Skat logs\ un reports\ mapes.
  pause
)
endlocal & exit /b %RC%
