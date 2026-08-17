@echo off
rem Forward to authoritative build.bat
call "%~dp0build.bat" %*
exit /b %ERRORLEVEL%
