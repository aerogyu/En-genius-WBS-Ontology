@echo off
rem Build the WBS ontology, checks, report and viewer from a WBS Excel file.
rem   Drag the WBS .xlsx onto this file  -> output\<time>\  report.xlsx, wbs_view.html, wbs.ttl, derived.ttl, run_log.json
rem   or: scripts\build.bat <file.xlsx> [--sheet NAME]
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
if "%~1"=="" goto :no_file
"%PY%" -m engenius_wbs build %*
set RC=%errorlevel%
echo.
if "%RC%"=="0" echo [build] RESULT: OK  (rc=0) - open wbs_view.html in the output folder printed above
if not "%RC%"=="0" echo [build] RESULT: rc=%RC%
pause
exit /b %RC%

:no_file
echo [NG] Drag the WBS .xlsx file onto build.bat
pause
exit /b 2

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
