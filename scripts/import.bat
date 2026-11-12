@echo off
rem Import a WBS Excel file into the tool database (data\real\wbs.db, one SQLite file). Drag the WBS .xlsx onto this file.
rem   Cells that differ from the original are listed in data\real\wbs.import_notes.json (stays in-house).
rem   If a database already exists it asks first; yes = old one renamed to .bak-<time> (kept), roles and Windows accounts carried over.
rem   Without asking: add --force
rem   or: scripts\import.bat <file.xlsx> [--db FILE] [--sheet NAME] [--force]
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
if "%~1"=="" goto :no_file
"%PY%" -m engenius_wbs import --ask %*
set RC=%errorlevel%
echo.
if "%RC%"=="0" echo [import] RESULT: OK  (rc=0)
if not "%RC%"=="0" echo [import] RESULT: rc=%RC%
pause
exit /b %RC%

:no_file
echo [NG] Drag the WBS .xlsx file onto import.bat
pause
exit /b 2

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
