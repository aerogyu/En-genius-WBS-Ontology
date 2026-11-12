@echo off
rem Apply what-if edits from the viewer to the WBS Excel -> a NEW Excel copy in the same format (the original is never changed).
rem   Select BOTH files (the original WBS .xlsx and the change list .json downloaded from tab 4 of wbs_view.html)
rem   and drag them onto this file together. Order does not matter.
rem   -> <original>_(edited)_v01.xlsx (next free number) next to the original + .changes.json, then build runs on the new file.
rem   or: scripts\apply.bat <file.xlsx> <changes.json> [--out DIR] [--no-build] [--force]
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
if "%~2"=="" goto :no_file
"%PY%" -m engenius_wbs apply %*
set RC=%errorlevel%
echo.
if "%RC%"=="0" echo [apply] RESULT: OK  (rc=0) - new Excel and its viewer are printed above
if not "%RC%"=="0" echo [apply] RESULT: rc=%RC%
pause
exit /b %RC%

:no_file
echo [NG] Drag BOTH files onto apply.bat together: the original WBS .xlsx and the change list .json
pause
exit /b 2

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
