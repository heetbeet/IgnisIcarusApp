@echo off
set "PYTHON=%~dp0\bin\python\python.exe"
if not exist "%PYTHON%" set "PYTHON=python"

echo using %PYTHON%
cd /d %~dp0
:loop
"%PYTHON%" ignis_controller.py %*
goto loop
