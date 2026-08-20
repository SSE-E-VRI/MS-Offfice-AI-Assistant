# PowerShell script to Uninstall AI Assistant Office Add-in
$ErrorActionPreference = "SilentlyContinue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  AI Assistant Office Add-in - Uninstall                   " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$apps = @("Word", "Excel", "PowerPoint")
$progIds = @("MistralAI.Addin", "MistralAI.Connect", "MistralAI.TaskPaneControl")
$clsids = @("{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}", "{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}")

# 1. Remove Office Add-in Registrations
Write-Host "`n[1/3] Removing Office Add-in registrations..." -ForegroundColor Yellow
foreach ($app in $apps) {
    foreach ($progId in $progIds) {
        $key = "HKCU:\Software\Microsoft\Office\$app\Addins\$progId"
        if (Test-Path $key) {
            Remove-Item -Path $key -Recurse -Force
        }
    }
}

# Remove versioned add-in keys (14.0/15.0/16.0) that install.ps1 also writes
$versions = @("14.0", "15.0", "16.0")
foreach ($ver in $versions) {
    foreach ($app in $apps) {
        foreach ($progId in $progIds) {
            $versionedKey = "HKCU:\Software\Microsoft\Office\$ver\$app\Addins\$progId"
            if (Test-Path $versionedKey) {
                Remove-Item -Path $versionedKey -Recurse -Force
                Write-Host "      Removed versioned key: $versionedKey" -ForegroundColor DarkGray
            }
        }
    }
}

# 2. Remove COM CLSIDs and ProgIDs
Write-Host "[2/3] Removing COM CLSIDs and ProgIDs..." -ForegroundColor Yellow
$roots = @("HKCU:\Software\Classes")
if (Test-Path "HKCU:\Software\Classes\Wow6432Node") {
    $roots += "HKCU:\Software\Classes\Wow6432Node"
}

foreach ($root in $roots) {
    foreach ($c in $clsids) {
        $k = "$root\CLSID\$c"
        if (Test-Path $k) {
            Remove-Item -Path $k -Recurse -Force
        }
    }

    foreach ($p in $progIds) {
        $pk = "$root\$p"
        if (Test-Path $pk) {
            Remove-Item -Path $pk -Recurse -Force
        }
    }
}

# 3. Clean Resiliency DoNotDisable entries
Write-Host "[3/3] Cleaning Resiliency entries..." -ForegroundColor Yellow
$versions = @("14.0", "15.0", "16.0")
foreach ($ver in $versions) {
    foreach ($app in $apps) {
        $dndKey = "HKCU:\Software\Microsoft\Office\$ver\$app\Resiliency\DoNotDisableAddinList"
        if (Test-Path $dndKey) {
            Remove-ItemProperty -Path $dndKey -Name "MistralAI.Addin" -ErrorAction SilentlyContinue
        }
    }
}

Write-Host "`n============================================================" -ForegroundColor Green
Write-Host "  [SUCCESS] AI Assistant Add-in completely uninstalled.     " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
exit 0
