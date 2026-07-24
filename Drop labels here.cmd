@echo off
rem Drag one or more Audacity label .txt files onto this file.
setlocal

if "%~1"=="" (
  echo.
  echo   Drag one or more Audacity label .txt files onto this file.
  echo.
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0AudacityLabelsToExcel.ps1" %*

if errorlevel 1 (
  echo.
  pause
)
exit /b %errorlevel%
