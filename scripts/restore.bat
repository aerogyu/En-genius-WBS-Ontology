@echo off
rem Put a backup copy back as the tool database (data\real\wbs.db). Drag a wbs_*.db file from data\real\backup onto this .bat.
rem   Close the server window first (restore refuses while the server is running).
rem   The current database is kept as wbs.db.bak-<time>. It asks before doing anything.
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
if "%~1"=="" goto :usage
"%PY%" -m engenius_wbs restore "%~1" %2 %3 %4
set RC=%errorlevel%
echo.
if "%RC%"=="0" echo [restore] RESULT: OK  (rc=0)
if not "%RC%"=="0" echo [restore] RESULT: rc=%RC%
pause
exit /b %RC%

:usage
echo Drag a backup file (data\real\backup\wbs_*.db) onto restore.bat
echo   or: scripts\restore.bat data\real\backup\wbs_20261113-180000_r12_daily.db
pause
exit /b 2

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
