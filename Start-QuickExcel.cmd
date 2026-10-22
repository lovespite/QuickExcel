@echo off
setlocal
cd /d "%~dp0"
if exist "artifacts\QuickExcel\QuickExcel.exe" (
  start "" "artifacts\QuickExcel\QuickExcel.exe" %*
  exit /b 0
)
dotnet run --project "QuickExcel.Desktop\QuickExcel.Desktop.csproj" --artifacts-path "%~dp0artifacts\build" -- %*
if errorlevel 1 pause
