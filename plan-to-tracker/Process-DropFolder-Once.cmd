@echo off
rem Processes whatever is in the drop folder now, then stops. Use it by hand, or from Task Scheduler.
rem   Process-DropFolder-Once.cmd [folder] [nopause]
setlocal
set "DROP=%~1"
if "%DROP%"=="" set "DROP=%~dp0DropHere"
if "%DROP:~-1%"=="\" set "DROP=%DROP:~0,-1%"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0plan_to_tracker.ps1" -Watch "%DROP%" -Once
if /i not "%~2"=="nopause" pause
