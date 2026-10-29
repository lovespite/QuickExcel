@echo off
setlocal
cd /d "%~dp0"
@REM  if exist "artifacts\QuickExcel\QuickExcel.exe" (
@REM    start "" "artifacts\QuickExcel\QuickExcel.exe" %*
@REM    exit /b 0
@REM  )
dotnet run --project "QuickExcel.Desktop\QuickExcel.Desktop.csproj" --artifacts-path "%~dp0artifacts\build" -- %*
if errorlevel 1 pause
