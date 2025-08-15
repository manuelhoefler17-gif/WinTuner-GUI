@echo off
setlocal

rem -- Get current directory
set "SCRIPT_DIR=%~dp0"

rem -- Path to PowerShell 7
set "PSH=%ProgramFiles%\PowerShell\7\pwsh.exe"

rem -- Check if PowerShell 7 exists
if not exist "%PSH%" (
    echo PowerShell 7 not found at %PSH%
    pause
    exit /b 1
)

rem -- Run the PowerShell script from the same directory
"%PSH%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%WinTuner_GUI.ps1"
