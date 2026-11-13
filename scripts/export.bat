@echo off
rem Export the tool database (data\real\wbs.db) to an Excel file in the original layout -> output\<time>\<name>_export.xlsx
rem   Person columns come out blank. A hidden UID column is added at the right end.
rem   or: scripts\export.bat [--db FILE] [--out FILE.xlsx]
rem   One module (keeps the full-project row numbers, adds an 'outside links' sheet): scripts\export.bat --module "<Lv1 name or WBS code>"
rem   One file per module: scripts\export.bat --by-module      A frozen snapshot: scripts\export.bat --snapshot <name or id>
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
"%PY%" -m engenius_wbs export %*
set RC=%errorlevel%
echo.
if "%RC%"=="0" echo [export] RESULT: OK  (rc=0)
if not "%RC%"=="0" echo [export] RESULT: rc=%RC%
pause
exit /b %RC%

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
