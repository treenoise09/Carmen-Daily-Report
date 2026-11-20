@echo off
cd /d "%~dp0scripts"
call npm.cmd start
if errorlevel 1 (
  echo Daily report failed. See the error above.
)
pause
