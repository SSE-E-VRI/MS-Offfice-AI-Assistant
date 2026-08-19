@echo off
setlocal EnableExtensions
cd /d "%~dp0"

rem Do not call bare "powershell" — Windows may open the "how do you want to open this" picker
rem (Notepad, Cursor, Antigravity) when .ps1 / powershell aliases are broken.
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" (
    echo [ERROR] Windows PowerShell was not found.
    echo Expected: %SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe
    pause
    exit /b 1
)

"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
    echo.
    echo [ERROR] Installation encountered errors. Exit code %EXITCODE%.
) else (
    echo.
)

pause
endlocal & exit /b %EXITCODE%
