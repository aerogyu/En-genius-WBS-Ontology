@echo off
rem Run the tests (synthetic data only).
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
setlocal
cd /d "%~dp0.."
call "%~dp0env.bat"
"%PY%" -m pytest -q
set RC=%errorlevel%
echo.
echo [selftest] rc=%RC%
pause
exit /b %RC%
