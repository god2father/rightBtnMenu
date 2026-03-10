@echo off
setlocal
set "SCRIPT_DIR=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%install-shortcuts.ps1"
set "EXIT_CODE=%ERRORLEVEL%"
if "%EXIT_CODE%"=="0" (
    echo.
    echo Shortcuts installed.
) else (
    echo.
    echo Shortcut install failed. Exit code: %EXIT_CODE%
)
pause
exit /b %EXIT_CODE%

