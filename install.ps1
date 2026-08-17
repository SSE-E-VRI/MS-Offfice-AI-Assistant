# PowerShell script to Build and Install AI Assistant Office Add-in
$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AI Assistant Office Add-in - Build & Install              " -ForegroundColor Cyan
Write-Host "  Targets: Word, Excel, PowerPoint (Office 2010 - 365)      " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $baseDir

# --- 1. Locate MSBuild ---
Write-Host "`n[1/5] Locating MSBuild..." -ForegroundColor Yellow
$msbuildCandidates = @(
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe",
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
)
$msbuild = $msbuildCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $msbuild) {
    Write-Error "MSBuild.exe not found under .NET Framework directory."
    exit 1
}
Write-Host "      Found MSBuild: $msbuild" -ForegroundColor DarkGray

# --- 2. Ensure NuGet CLI and Restore Packages ---
Write-Host "`n[2/5] Restoring NuGet Packages..." -ForegroundColor Yellow
$nugetPath = Join-Path $baseDir "nuget.exe"
if (-not (Test-Path $nugetPath)) {
    Write-Host "      Downloading nuget.exe..." -ForegroundColor DarkGray
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    (New-Object Net.WebClient).DownloadFile('https://dist.nuget.org/win-x86-commandline/latest/nuget.exe', $nugetPath)
}

& $nugetPath restore "$baseDir\src\packages.config" -PackagesDirectory "$baseDir\packages" | Out-Null
Write-Host "      Packages restored." -ForegroundColor DarkGray

# --- 3. Build x86 and x64 Release Configurations ---
Write-Host "`n[3/5] Building x86 & x64 Release Configurations..." -ForegroundColor Yellow
$projPath = Join-Path $baseDir "src\MistralOfficeAddin.csproj"

& $msbuild $projPath /p:Configuration=Release /p:Platform=x86 /v:m /nologo
if ($LASTEXITCODE -ne 0) {
    Write-Error "x86 Build failed."
    exit 1
}

& $msbuild $projPath /p:Configuration=Release /p:Platform=x64 /v:m /nologo
if ($LASTEXITCODE -ne 0) {
    Write-Error "x64 Build failed."
    exit 1
}
Write-Host "      Built x86 and x64 Release assemblies." -ForegroundColor DarkGray

# --- 4. Register COM Classes (Current User) ---
Write-Host "`n[4/5] Registering COM classes for Current User..." -ForegroundColor Yellow
$dll64 = Join-Path $baseDir "bin\x64\Release\MistralOfficeAddin.dll"
$dll32 = Join-Path $baseDir "bin\x86\Release\MistralOfficeAddin.dll"
$reg64 = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$reg32 = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"

$guidConnect = "{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}"
$guidTaskPane = "{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}"

# 64-bit COM Registration
if (Test-Path $dll64) {
    $regFile64 = "$env:TEMP\MistralAI64.reg"
    & $reg64 $dll64 /nologo /codebase /regfile:$regFile64 | Out-Null
    if (Test-Path $regFile64) {
        $content = Get-Content -Raw $regFile64
        $content = $content -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes"
        Set-Content -Path $regFile64 -Value $content -Encoding Unicode
        reg import $regFile64 | Out-Null
        Remove-Item -Path $regFile64 -Force -ErrorAction SilentlyContinue
    }
}

# 32-bit COM Registration (for 32-bit Office)
if (Test-Path $dll32) {
    $regFile32 = "$env:TEMP\MistralAI32.reg"
    & $reg32 $dll32 /nologo /codebase /regfile:$regFile32 | Out-Null
    if (Test-Path $regFile32) {
        $content = Get-Content -Raw $regFile32
        $content = $content -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes\Wow6432Node"
        Set-Content -Path $regFile32 -Value $content -Encoding Unicode
        reg import $regFile32 | Out-Null
        Remove-Item -Path $regFile32 -Force -ErrorAction SilentlyContinue
    }
}

# Setup CodeBase & ActiveX Task Pane Categories
$roots = @("HKCU:\Software\Classes")
if (Test-Path "HKCU:\Software\Classes\Wow6432Node") {
    $roots += "HKCU:\Software\Classes\Wow6432Node"
}

$u64 = "file:///" + ($dll64 -replace "\\", "/")
$u32 = "file:///" + ($dll32 -replace "\\", "/")
$clsids = @($guidConnect, $guidTaskPane)

foreach ($root in $roots) {
    $is32 = $root -like "*Wow6432Node*"
    $codebase = if ($is32) { $u32 } else { $u64 }

    foreach ($c in $clsids) {
        $inproc = "$root\CLSID\$c\InprocServer32"
        if (Test-Path $inproc) {
            Set-ItemProperty -Path $inproc -Name "CodeBase" -Value $codebase -ErrorAction SilentlyContinue
        }
    }

    $k = "$root\CLSID\$guidTaskPane"
    if (Test-Path $k) {
        New-Item -Path "$k\Control" -Force | Out-Null
        New-Item -Path "$k\MiscStatus" -Force | Out-Null
        Set-ItemProperty -Path "$k\MiscStatus" -Name "(default)" -Value "131473" -ErrorAction SilentlyContinue
        $cat = "$k\Implemented Categories"
        New-Item -Path "$cat\{7DD95801-9882-11CF-9FA9-00AA006C42C4}" -Force | Out-Null
        New-Item -Path "$cat\{7DD95802-9882-11CF-9FA9-00AA006C42C4}" -Force | Out-Null
        New-Item -Path "$cat\{40FC6ED4-2438-11CF-A3DB-080036F12502}" -Force | Out-Null
    }
}

# Optional mirror for Word 2010 ActiveX docking (if running elevated)
try {
    $hkcuClsid = "HKCU\Software\Classes\CLSID\$guidTaskPane"
    $hklmClsid = "HKLM\Software\Classes\CLSID\$guidTaskPane"
    $hkcuProg = "HKCU\Software\Classes\MistralAI.TaskPaneControl"
    $hklmProg = "HKLM\Software\Classes\MistralAI.TaskPaneControl"
    Start-Process reg.exe -ArgumentList "copy `"$hkcuClsid`" `"$hklmClsid`" /s /f" -NoNewWindow -Wait -ErrorAction SilentlyContinue
    Start-Process reg.exe -ArgumentList "copy `"$hkcuProg`" `"$hklmProg`" /s /f" -NoNewWindow -Wait -ErrorAction SilentlyContinue
} catch { }

# --- 5. Register Office Addin Keys & Enable ---
Write-Host "`n[5/5] Configuring Office Add-in Registration..." -ForegroundColor Yellow
$apps = @("Word", "Excel", "PowerPoint")
$versions = @("14.0", "15.0", "16.0")

foreach ($ver in $versions) {
    foreach ($app in $apps) {
        $resiliencyBase = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency"
        if (Test-Path $resiliencyBase) {
            $doNotDisable = "$resiliencyBase\DoNotDisableAddinList"
            if (-not (Test-Path $doNotDisable)) {
                New-Item -Path $doNotDisable -Force | Out-Null
            }
            if ((Get-ItemProperty -Path $doNotDisable -Name "MistralAI.Addin" -ErrorAction SilentlyContinue) -eq $null) {
                Set-ItemProperty -Path $doNotDisable -Name "MistralAI.Addin" -Value 1 -Type DWord -ErrorAction SilentlyContinue
            }
        }
    }
}

foreach ($app in $apps) {
    $oldKey = "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Connect"
    if (Test-Path $oldKey) {
        Remove-Item -Path $oldKey -Recurse -Force -ErrorAction SilentlyContinue
    }

    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Addin"
    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    Set-ItemProperty -Path $key -Name "FriendlyName" -Value "AI Assistant"
    Set-ItemProperty -Path $key -Name "Description" -Value "AI Assistant for Word, Excel, and PowerPoint"
    Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  [SUCCESS] AI Assistant Add-in built & installed!          " -ForegroundColor Green
Write-Host "  Launch Word, Excel, or PowerPoint to start using it.      " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
