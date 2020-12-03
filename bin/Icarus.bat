@echo off
cd /d "%~dp0.."
:loop
    python "%~dp0..\ignis_controller_icarus.py" %*
goto loop

