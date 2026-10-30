@echo off
setlocal
cd /d "%~dp0"
rem Per-process WPF software rendering; applies only to this launch.
dotnet run --project "QuickExcel.Desktop\QuickExcel.Desktop.csproj" -c Release --artifacts-path "%~dp0artifacts\build" -- --software-rendering %*
if errorlevel 1 pause
