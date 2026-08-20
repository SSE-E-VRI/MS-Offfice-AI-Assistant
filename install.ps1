# PowerShell script to Build and Install MS Office AI Assistant
$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  MS Office AI Assistant - Build & Install                  " -ForegroundColor Cyan
Write-Host "  Targets: Word, Excel, PowerPoint (Office 2010 - 365)      " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $baseDir

function Import-RegFileSafe([string]$regFilePath) {
    if ([string]::IsNullOrWhiteSpace($regFilePath) -or -not (Test-Path $regFilePath)) { return }
    $regExe = Join-Path $env:WINDIR "System32\reg.exe"
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    try {
        & $regExe import $regFilePath 2>$null | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Host "      Registry import warning (non-fatal, often HKLM access denied)." -ForegroundColor DarkGray
        }
    } catch {
        Write-Host "      Registry import skipped (non-fatal)." -ForegroundColor DarkGray
    } finally {
        $ErrorActionPreference = $prev
        $global:LASTEXITCODE = 0
    }
}

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
$projPath = Join-Path $baseDir "src\MSOfficeAIAssistant.csproj"

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
$dll64 = Join-Path $baseDir "bin\x64\Release\MSOfficeAIAssistant.dll"
$dll32 = Join-Path $baseDir "bin\x86\Release\MSOfficeAIAssistant.dll"
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
        Import-RegFileSafe $regFile64
        Remove-Item -Path $regFile64 -Force -ErrorAction SilentlyContinue
    }
}

# 32-bit COM Registration (for 32-bit Office)
# IMPORTANT: On 64-bit Windows, never import the 32-bit regfile into
# HKCU\Software\Classes (non-Wow). That overwrites the 64-bit CLSID and
# leaves InprocServer32\<version>\CodeBase pointing at the x86 DLL, which
# makes 64-bit Excel/Word fail with 0x8007000B (Bad Image Format).
if (Test-Path $dll32) {
    $regFile32 = "$env:TEMP\MistralAI32.reg"
    & $reg32 $dll32 /nologo /codebase /regfile:$regFile32 | Out-Null
    if (Test-Path $regFile32) {
        $content = Get-Content -Raw $regFile32
        $is64BitOs = [Environment]::Is64BitOperatingSystem

        if (-not $is64BitOs) {
            # Pure 32-bit Windows: Classes is not redirected.
            $contentDirect = $content -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes"
            Set-Content -Path $regFile32 -Value $contentDirect -Encoding Unicode
            Import-RegFileSafe $regFile32
        } else {
            # 32-bit Office on 64-bit Windows reads Wow6432Node.
            $contentWow = $content -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes\Wow6432Node"
            Set-Content -Path $regFile32 -Value $contentWow -Encoding Unicode
            Import-RegFileSafe $regFile32
        }

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
            # Parent + versioned subkeys (CLR prefers InprocServer32\<version>\CodeBase).
            $inprocKeys = @($inproc)
            Get-ChildItem -Path $inproc -ErrorAction SilentlyContinue | ForEach-Object {
                $inprocKeys += $_.PSPath
            }
            foreach ($ik in $inprocKeys) {
                Set-ItemProperty -Path $ik -Name "CodeBase" -Value $codebase -ErrorAction SilentlyContinue
            }
        }
    }

    $k = "$root\CLSID\$guidTaskPane"
    if (Test-Path $k) {
        try {
            New-Item -Path "$k\Control" -Force -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$k\MiscStatus" -Force -ErrorAction SilentlyContinue | Out-Null
            Set-ItemProperty -Path "$k\MiscStatus" -Name "(default)" -Value "131473" -ErrorAction SilentlyContinue
            $cat = "$k\Implemented Categories"
            New-Item -Path "$cat\{7DD95801-9882-11CF-9FA9-00AA006C42C4}" -Force -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$cat\{7DD95802-9882-11CF-9FA9-00AA006C42C4}" -Force -ErrorAction SilentlyContinue | Out-Null
            New-Item -Path "$cat\{40FC6ED4-2438-11CF-A3DB-080036F12502}" -Force -ErrorAction SilentlyContinue | Out-Null
        } catch {
            Write-Host "      Skipped ActiveX category write under $root (non-fatal)." -ForegroundColor DarkGray
        }
    }
}

# Optional mirror for Word 2010 ActiveX docking (if running elevated). Access denied is expected without admin.
try {
    $regExe = Join-Path $env:WINDIR "System32\reg.exe"
    $hkcuClsid = "HKCU\Software\Classes\CLSID\$guidTaskPane"
    $hklmClsid = "HKLM\Software\Classes\CLSID\$guidTaskPane"
    $hkcuProg = "HKCU\Software\Classes\MistralAI.TaskPaneControl"
    $hklmProg = "HKLM\Software\Classes\MistralAI.TaskPaneControl"
    $prev = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    & $regExe copy $hkcuClsid $hklmClsid /s /f 2>$null | Out-Null
    $copy1 = $LASTEXITCODE
    & $regExe copy $hkcuProg $hklmProg /s /f 2>$null | Out-Null
    $copy2 = $LASTEXITCODE
    $ErrorActionPreference = $prev
    $global:LASTEXITCODE = 0
    if ($copy1 -ne 0 -or $copy2 -ne 0) {
        Write-Host "      HKLM ActiveX mirror skipped (admin usually required)." -ForegroundColor DarkGray
    }
} catch {
    Write-Host "      HKLM ActiveX mirror skipped (admin usually required)." -ForegroundColor DarkGray
    $global:LASTEXITCODE = 0
}

# --- 5. Register Office Addin Keys & Enable ---
Write-Host "`n[5/5] Configuring Office Add-in Registration..." -ForegroundColor Yellow
$apps = @("Word", "Excel", "PowerPoint")
$versions = @("14.0", "15.0", "16.0")

foreach ($ver in $versions) {
    foreach ($app in $apps) {
        try {
            $resiliencyBase = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency"
            if (Test-Path $resiliencyBase) {
                $doNotDisable = "$resiliencyBase\DoNotDisableAddinList"
                if (-not (Test-Path $doNotDisable)) {
                    New-Item -Path $doNotDisable -Force -ErrorAction SilentlyContinue | Out-Null
                }
                if (Test-Path $doNotDisable) {
                    Set-ItemProperty -Path $doNotDisable -Name "MistralAI.Addin" -Value 1 -Type DWord -ErrorAction SilentlyContinue
                }
            }
        } catch {
            Write-Host "      Skipped resiliency key for $app $ver (access denied is non-fatal)." -ForegroundColor DarkGray
        }
    }
}

$addinRegistered = $false
foreach ($app in $apps) {
    try {
        $oldKey = "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Connect"
        if (Test-Path $oldKey) {
            Remove-Item -Path $oldKey -Recurse -Force -ErrorAction SilentlyContinue
        }

        $keys = @(
            "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Addin"
        )
        foreach ($ver in $versions) {
            $keys += "HKCU:\Software\Microsoft\Office\$ver\$app\Addins\MistralAI.Addin"
        }

        foreach ($key in $keys) {
            if (-not (Test-Path $key)) {
                New-Item -Path $key -Force -ErrorAction SilentlyContinue | Out-Null
            }
            if (Test-Path $key) {
                Set-ItemProperty -Path $key -Name "FriendlyName" -Value "MS Office AI Assistant" -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $key -Name "Description" -Value "MS Office AI Assistant for Word, Excel, and PowerPoint" -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord -ErrorAction SilentlyContinue
                Set-ItemProperty -Path $key -Name "CommandLineSafe" -Value 0 -Type DWord -ErrorAction SilentlyContinue
            }
        }
        $addinRegistered = $true
        Write-Host "      Registered $app add-in (HKCU, LoadBehavior=3)." -ForegroundColor DarkGray
    } catch {
        Write-Host "      WARNING: Could not register $app add-in: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Remove only our add-in from resiliency lists (do NOT wipe other add-ins' entries)
foreach ($ver in $versions) {
    foreach ($app in $apps) {
        try {
            $disabled = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency\DisabledItems"
            if (Test-Path $disabled) {
                Remove-ItemProperty -Path $disabled -Name "MistralAI.Addin" -ErrorAction SilentlyContinue
            }
            $crash = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency\CrashingAddinList"
            if (Test-Path $crash) {
                Remove-ItemProperty -Path $crash -Name "MistralAI.Addin" -ErrorAction SilentlyContinue
            }
        } catch { }
    }
}

if (-not $addinRegistered) {
    Write-Host "      ERROR: No Office add-in keys could be written. Close Word/Excel/PowerPoint and retry." -ForegroundColor Red
    exit 1
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  [SUCCESS] MS Office AI Assistant built & installed!       " -ForegroundColor Green
Write-Host "  Launch Word, Excel, or PowerPoint to start using it.      " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
# Do not leak a failed native "reg import" exit code to cmd.exe.
exit 0
