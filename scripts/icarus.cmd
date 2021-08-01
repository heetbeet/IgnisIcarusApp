@echo off
:loop
    call "%~dp0python.cmd" "%~dp0../src/ignis_controller_icarus.py" %*
goto loop

