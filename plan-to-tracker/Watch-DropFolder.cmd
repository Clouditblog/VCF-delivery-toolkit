@echo off
rem Leave this window open. Drop a Word plan (.docx) into the DropHere folder and the tracker
rem appears in DropHere\Trackers. The Word file moves to DropHere\Processed (or DropHere\Failed).
rem To watch a different folder:  Watch-DropFolder.cmd "D:\Plans\DropHere"
setlocal
set "DROP=%~1"
if "%DROP%"=="" set "DROP=%~dp0DropHere"
if "%DROP:~-1%"=="\" set "DROP=%DROP:~0,-1%"
echo Drop folder: %DROP%
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0plan_to_tracker.ps1" -Watch "%DROP%"
echo.
pause
