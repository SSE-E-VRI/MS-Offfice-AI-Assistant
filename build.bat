@echo off
setlocal enabledelayedexpansion

echo ===================================================================
echo   Mistral AI Office Add-in - Cross-Version Build Script
echo   Targets: Word, Excel, PowerPoint, Outlook (Office 2010 - 365)
echo   Architectures: x86 (32-bit) and x64 (64-bit)
echo ===================================================================
echo.

cd /d "%~dp0"

:: 1. Locate MSBuild
set MSBUILD=
if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe" (
    set "MSBUILD=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
) else if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" (
    set "MSBUILD=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
)

if "%MSBUILD%"=="" (
    echo [ERROR] MSBuild.exe not found under .NET Framework directory.
    exit /b 1
)
echo [OK] Using MSBuild: %MSBUILD%

:: 2. Ensure NuGet CLI exists
if not exist "nuget.exe" (
    echo [INFO] Downloading nuget.exe...
    powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; (New-Object Net.WebClient).DownloadFile('https://dist.nuget.org/win-x86-commandline/latest/nuget.exe', 'nuget.exe')"
    if errorlevel 1 (
        echo [ERROR] Failed to download nuget.exe.
        exit /b 1
    )
)
echo [OK] NuGet CLI is ready.

:: 3. Restore Packages
echo.
echo [1/3] Restoring NuGet Packages...
.\nuget.exe restore src\packages.config -PackagesDirectory packages
if errorlevel 1 (
    echo [ERROR] NuGet package restore failed.
    exit /b 1
)

:: 4. Build x86 (32-bit) Release
echo.
echo [2/3] Building x86 (32-bit) Release Configuration...
"%MSBUILD%" src\MistralOfficeAddin.csproj /p:Configuration=Release /p:Platform=x86 /v:m
if errorlevel 1 (
    echo [ERROR] x86 Build failed.
    exit /b 1
)
echo [OK] x86 Output: bin\x86\Release\MistralOfficeAddin.dll

:: 5. Build x64 (64-bit) Release
echo.
echo [3/3] Building x64 (64-bit) Release Configuration...
"%MSBUILD%" src\MistralOfficeAddin.csproj /p:Configuration=Release /p:Platform=x64 /v:m
if errorlevel 1 (
    echo [ERROR] x64 Build failed.
    exit /b 1
)
echo [OK] x64 Output: bin\x64\Release\MistralOfficeAddin.dll

:: 6. Optional Local Registration
set REGASM32=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\regasm.exe
set REGASM64=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe

echo.
echo ===================================================================
echo [SUCCESS] Both x86 and x64 assemblies built successfully!
echo.
echo To register for local testing:
echo   - 32-bit Office: "%REGASM32%" /codebase bin\x86\Release\MistralOfficeAddin.dll
echo   - 64-bit Office: "%REGASM64%" /codebase bin\x64\Release\MistralOfficeAddin.dll
echo ===================================================================

endlocal
exit /b 0
