@echo off
setlocal
cd /d "%~dp0"
rem Per-launch Avalonia software rendering for graphics diagnosis.
set "QUICKEXCEL_SOFTWARE_RENDERING=1"
call "%~dp0Start-QuickExcel.cmd" %*
