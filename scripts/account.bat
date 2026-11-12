@echo off
rem Windows account -> role for the WBS server (v0.5 S2). Run from a command prompt; double-click only lists them.
rem   scripts\account.bat set DOMAIN\user --role admin1     (role login ID from: scripts\role.bat list)
rem   scripts\account.bat remove DOMAIN\user
rem   Accounts not in the list come in as the shared viewer role (config windows_login.unregistered: viewer | deny).
rem   The list stays in the server database only; history shows role names only.
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
if not exist "%PY%" goto :no_py
if "%~1"=="" goto :list
"%PY%" -m engenius_wbs account %*
exit /b %errorlevel%

:list
"%PY%" -m engenius_wbs account list
echo.
echo Add (command prompt): scripts\account.bat set DOMAIN\user --role ROLE_LOGIN_ID
pause
exit /b 0

:no_py
echo [NG] Python not found: %PY%
echo      Set TOOL_ONTO_DIR in scripts\env.bat to the Tool ontology folder (it has .venv).
pause
exit /b 2
