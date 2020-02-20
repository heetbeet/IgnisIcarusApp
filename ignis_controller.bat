
''' >NUL  2>NUL
@echo off
cd /d %~dp0
:loop
python ignis_controller.py %*
goto loop
'''
