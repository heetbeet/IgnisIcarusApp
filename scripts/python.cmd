@echo off
setlocal

set "py=%~dp0..\bin\python\python.exe"
if not exist "%py%" set "py=python"

call "%py%"  %*
exit /b %errorlevel%


