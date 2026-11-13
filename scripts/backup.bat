@echo off
rem Make a copy of the tool database now (safe while the server is running) -> data\real\backup\wbs_<time>_r<rev>_manual.db
rem   The server also backs up by itself: when it starts, once a day (backup.daily_at, default 18:00) and right before a snapshot.
rem   List copies: scripts\backup.bat list      Other folder: config\wbs.yaml  backup: dir: D:\wbs_backup
rem   Copies include the Windows account -> role table: keep the backup folder as safe as the server PC.
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
"%PY%" -m engenius_wbs backup %*
set RC=%errorlevel%
echo.
if "%RC%"=="0" echo [backup] RESULT: OK  (rc=0)
if not "%RC%"=="0" echo [backup] RESULT: rc=%RC%
pause
exit /b %RC%

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
