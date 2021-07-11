
''' >NUL  2>NUL
@echo off
cd /d %~dp0
:loop
"%~dp0\bin\python\python.exe" ignis_controller.py %*
goto loop
'''
