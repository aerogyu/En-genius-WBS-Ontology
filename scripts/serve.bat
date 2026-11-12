@echo off
rem Start the WBS read-only server (v0.5 S2) for the whole in-house network: http://<this PC IP>:8100/
rem   Keep this window open while people use it; Ctrl+C stops it. Uses data\real\wbs.db.
rem   First time: Windows may ask to allow network access (firewall) - that is your decision (private network only).
rem   Needs at least one admin role: scripts\role.bat add admin1 --label "Admin 1" --kind admin
rem   Other port: scripts\serve.bat --port 8200    Only this PC: scripts\serve.bat --host 127.0.0.1
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
"%PY%" -m engenius_wbs serve --host 0.0.0.0 %*
set RC=%errorlevel%
echo.
echo [serve] stopped (rc=%RC%)
pause
exit /b %RC%

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
