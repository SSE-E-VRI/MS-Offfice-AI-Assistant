# PowerShell script to unregister AI Assistant Office Add-in
$ErrorActionPreference = "SilentlyContinue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Unregistering AI Assistant Office Add-in                  " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$apps = @("Word", "Excel", "PowerPoint")
$progIds = @("MistralAI.Addin", "MistralAI.Connect")

foreach ($app in $apps) {
    foreach ($progId in $progIds) {
        $key = "HKCU:\Software\Microsoft\Office\$app\Addins\$progId"
        if (Test-Path $key) {
            Remove-Item -Path $key -Recurse -Force
        }
    }
}

$roots = @("HKCU:\Software\Classes")
if (Test-Path "HKCU:\Software\Classes\Wow6432Node") {
    $roots += "HKCU:\Software\Classes\Wow6432Node"
}

$clsids = @("{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}", "{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}")

foreach ($root in $roots) {
    foreach ($c in $clsids) {
        $k = "$root\CLSID\$c"
        if (Test-Path $k) {
            Remove-Item -Path $k -Recurse -Force
        }
    }
}

Write-Host "`n[SUCCESS] Add-in unregistered." -ForegroundColor Green
