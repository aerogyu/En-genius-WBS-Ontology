@echo off
rem Local settings. Every other script calls this file first (so it has no pause of its own).
rem This .bat is ASCII-only on purpose: cmd reads .bat files as cp949 on Korean Windows.
rem No venv of its own: this repo uses the Python of the sibling Tool ontology repo (openpyxl, rdflib, pyyaml, pytest are there).
if "%TOOL_ONTO_DIR%"=="" set "TOOL_ONTO_DIR=%~dp0..\..\engenius-tool-ontology"
set "PY=%TOOL_ONTO_DIR%\.venv\Scripts\python.exe"
set "PYTHONPATH=%~dp0..\src"
rem Python UTF-8 mode: config and output are UTF-8 with Korean text (same as the sibling repos).
set PYTHONUTF8=1
