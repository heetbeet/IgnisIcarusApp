@echo off
pushd "%~dp0"
    call "%~dp0..\scripts\python.cmd" test_doctests.py
popd

if /i "%comspec% /c ``%~0` `" equ "%cmdcmdline:"=`%" pause