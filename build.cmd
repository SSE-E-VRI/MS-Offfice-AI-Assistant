@echo off
rem ============================================================
rem  Builds MistralOfficeAddin.dll with the .NET Framework 4.x
rem  C# compiler - no Visual Studio required.
rem  Target: .NET Framework 4.0 (runs on 4.0 through 4.8, so it
rem  works on Office 2010-era machines as well as Windows 11).
rem ============================================================
setlocal
cd /d "%~dp0"

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" (
    echo ERROR: .NET Framework 4.x csc.exe not found. Install .NET Framework 4.x first.
    exit /b 1
)

rem Find sn.exe for strong naming
set SN=
for /d %%d in ("%ProgramFiles(x86)%\Microsoft SDKs\Windows\*") do (
    if exist "%%d\bin\NETFX 4.8 Tools\sn.exe" set SN=%%d\bin\NETFX 4.8 Tools\sn.exe
    if exist "%%d\bin\NETFX 4.6.2 Tools\sn.exe" set SN=%%d\bin\NETFX 4.6.2 Tools\sn.exe
    if exist "%%d\bin\NETFX 4.6.1 Tools\sn.exe" set SN=%%d\bin\NETFX 4.6.1 Tools\sn.exe
)
if "%SN%"=="" set SN=sn.exe

if not exist bin mkdir bin

set KEYOPT=
if not exist MistralAI.snk (
    echo Generating strong name key...
    "%SN%" -q -k MistralAI.snk >nul 2>&1
    if errorlevel 1 (
        echo WARNING: sn.exe not found in standard paths. Assembly will not be strong-name signed.
    ) else (
        echo Key generated.
    )
)
if exist MistralAI.snk set KEYOPT=/keyfile:MistralAI.snk

"%CSC%" /nologo /target:library /platform:anycpu /optimize+ %KEYOPT% /warnaserror ^
  /out:bin\MistralOfficeAddin.dll ^
  /r:System.dll /r:System.Core.dll /r:System.Drawing.dll ^
  /r:System.Windows.Forms.dll /r:System.Web.Extensions.dll ^
  /r:System.Security.dll /r:Microsoft.CSharp.dll ^
  src\AssemblyInfo.cs src\ComInterfaces.cs src\RibbonXml.cs ^
  src\SettingsStore.cs src\MistralClient.cs src\Connect.cs ^
  src\TaskPaneControl.cs src\SettingsForm.cs

if errorlevel 1 (
    echo BUILD FAILED
    exit /b 1
)
echo Build OK: bin\MistralOfficeAddin.dll
endlocal
