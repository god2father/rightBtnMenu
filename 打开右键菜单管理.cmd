@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%dist\Right Click Menu Manager.exe"
if exist "%EXE_PATH%" (
    start "" "%EXE_PATH%"
    exit /b 0
)
start "" powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -STA -File "%SCRIPT_DIR%rightBtnMenu.ps1"
exit /b 0

