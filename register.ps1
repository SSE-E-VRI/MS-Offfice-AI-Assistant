# PowerShell script to register Mistral Office Add-in for current user
$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Registering Mistral Office Add-in (Current User)          " -ForegroundColor Cyan
Write-Host "  Covers: Word, Excel, PowerPoint, Outlook (2010 - 365)     " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dll64 = Join-Path $baseDir "bin\x64\Release\MistralOfficeAddin.dll"
$dll32 = Join-Path $baseDir "bin\x86\Release\MistralOfficeAddin.dll"

$reg64 = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$reg32 = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"

$guidConnect = "{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}"
$guidTaskPane = "{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}"

# 1. 64-bit COM Registration
if (Test-Path $dll64) {
    Write-Host "[1/4] Registering 64-bit COM classes..." -ForegroundColor Yellow
    $regFile64 = "$env:TEMP\MistralAI64.reg"
    & $reg64 $dll64 /nologo /codebase /regfile:$regFile64 | Out-Null
    if (Test-Path $regFile64) {
        $content = Get-Content -Raw $regFile64
        $content = $content -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes"
        Set-Content -Path $regFile64 -Value $content -Encoding Unicode
        reg import $regFile64 | Out-Null
    }
}

# 2. 32-bit COM Registration (for 32-bit Office)
if (Test-Path $dll32) {
    Write-Host "[2/4] Registering 32-bit COM classes (Wow6432Node)..." -ForegroundColor Yellow
    $regFile32 = "$env:TEMP\MistralAI32.reg"
    & $reg32 $dll32 /nologo /codebase /regfile:$regFile32 | Out-Null
    if (Test-Path $regFile32) {
        $content = Get-Content -Raw $regFile32
        $content = $content -replace "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\Software\Classes\Wow6432Node"
        Set-Content -Path $regFile32 -Value $content -Encoding Unicode
        reg import $regFile32 | Out-Null
    }
}

# 3. Setup CodeBase & ActiveX Task Pane Categories
Write-Host "[3/4] Configuring CodeBase and ActiveX safety categories..." -ForegroundColor Yellow
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

    # ActiveX categories for TaskPaneControl
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

# Office 2010 CreateCTP looks up ActiveX controls in HKLM, not per-user HKCU.
Write-Host "      Mirroring TaskPaneControl ActiveX registration to HKLM (required to dock in Word)..." -ForegroundColor Yellow
$hkcuClsid = "HKCU\Software\Classes\CLSID\$guidTaskPane"
$hklmClsid = "HKLM\Software\Classes\CLSID\$guidTaskPane"
$hkcuProg = "HKCU\Software\Classes\MistralAI.TaskPaneControl"
$hklmProg = "HKLM\Software\Classes\MistralAI.TaskPaneControl"
$copied = $false
cmd /c "reg copy `"$hkcuClsid`" `"$hklmClsid`" /s /f" | Out-Null
if ($LASTEXITCODE -eq 0) { $copied = $true }
cmd /c "reg copy `"$hkcuProg`" `"$hklmProg`" /s /f" | Out-Null
if ($copied) {
    Write-Host "      HKLM ActiveX registration complete." -ForegroundColor Green
} else {
    Write-Host "      WARNING: Could not write HKLM. Re-run this script from an elevated PowerShell so Word can dock the pane." -ForegroundColor Red
}

# 4. Clear Resiliency / Disabled Items & Enable Add-in in Office
Write-Host "[4/4] Clearing DisabledItems and enabling Add-in in Word, Excel, PowerPoint, Outlook..." -ForegroundColor Yellow
$apps = @("Word", "Excel", "PowerPoint", "Outlook")
$versions = @("14.0", "15.0", "16.0")

foreach ($ver in $versions) {
    foreach ($app in $apps) {
        $resiliencyBase = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency"
        if (Test-Path $resiliencyBase) {
            # 1. Delete all properties directly inside DisabledItems
            $disabledPath = "$resiliencyBase\DisabledItems"
            if (Test-Path $disabledPath) {
                Remove-Item -Path $disabledPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "  Cleared DisabledItems for $app ($ver)" -ForegroundColor Green
            }

            # 2. Delete CrashingAddinList and StartupItems
            foreach ($sub in @("CrashingAddinList", "StartupItems", "NotificationItems")) {
                $target = "$resiliencyBase\$sub"
                if (Test-Path $target) {
                    Remove-Item -Path $target -Recurse -Force -ErrorAction SilentlyContinue
                }
            }

            # 3. Add to DoNotDisableAddinList so Office never prompts to disable
            $doNotDisable = "$resiliencyBase\DoNotDisableAddinList"
            if (-not (Test-Path $doNotDisable)) {
                New-Item -Path $doNotDisable -Force | Out-Null
            }
            Set-ItemProperty -Path $doNotDisable -Name "MistralAI.Addin" -Value 1 -Type DWord -ErrorAction SilentlyContinue
        }
    }
}

# Clean old duplicate Connect progid
foreach ($app in $apps) {
    $oldKey = "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Connect"
    if (Test-Path $oldKey) {
        Remove-Item -Path $oldKey -Recurse -Force -ErrorAction SilentlyContinue
    }

    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Addin"
    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    Set-ItemProperty -Path $key -Name "FriendlyName" -Value "Mistral AI Assistant"
    Set-ItemProperty -Path $key -Name "Description" -Value "AI assistant using your own Mistral API key"
    Set-ItemProperty -Path $key -Name "LoadBehavior" -Value 3 -Type DWord
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  [SUCCESS] Mistral AI Add-in registered and enabled!       " -ForegroundColor Green
Write-Host "  Launch Word, Excel, PowerPoint, or Outlook to start.      " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
