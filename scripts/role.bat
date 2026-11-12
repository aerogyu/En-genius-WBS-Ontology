@echo off
rem Role accounts for the WBS server (v0.5 S2). Run from a command prompt; double-click only lists the roles.
rem   scripts\role.bat add admin1 --label "Admin 1" --kind admin
rem   scripts\role.bat add fan --label "FAN leader" --kind leader --module <Lv1 name or WBS code>
rem   scripts\role.bat add view --label "Viewer" --kind viewer
rem   scripts\role.bat passwd fan  /  off fan  /  on fan  /  list
rem   Labels are what the history shows: use role names, never a person's name.
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
if "%~1"=="" goto :list
"%PY%" -m engenius_wbs role %*
exit /b %errorlevel%

:list
"%PY%" -m engenius_wbs role list
echo.
echo Add a role (command prompt): scripts\role.bat add LOGIN --label "NAME" --kind admin^|leader^|viewer [--module LV1]
pause
exit /b 0

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
