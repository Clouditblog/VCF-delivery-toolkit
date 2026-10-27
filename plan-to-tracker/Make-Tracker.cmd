@echo off
rem Drag a Word plan (.docx) onto this file, or run:  Make-Tracker.cmd "My Plan.docx"
rem The tracker is written next to the Word file. An existing tracker is never overwritten.
setlocal
if "%~1"=="" (
    echo Drag a Word plan ^(.docx^) onto this file, or run:
    echo     Make-Tracker.cmd "My Plan.docx"
    echo.
    pause
    exit /b 1
)
:next
if "%~1"=="" goto done
echo.
echo Building a tracker from "%~1"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0plan_to_tracker.ps1" -Docx "%~1"
shift
goto next
:done
echo.
pause
