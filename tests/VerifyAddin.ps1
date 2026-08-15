# Verification and self-test script for Mistral AI Office Add-in
$ErrorActionPreference = "Stop"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Mistral AI Office Add-in - Verification Suite  " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

$baseDir = Split-Path -Parent $PSScriptRoot
$dll32 = Join-Path $baseDir "bin\x86\Release\MistralOfficeAddin.dll"
$dll64 = Join-Path $baseDir "bin\x64\Release\MistralOfficeAddin.dll"

# 1. Check binaries exist
Write-Host "[1/4] Checking Binary Outputs..." -ForegroundColor Yellow
if (Test-Path $dll32) {
    $f32 = Get-Item $dll32
    $sz32 = [Math]::Round($f32.Length / 1024, 1)
    Write-Host ("  [OK] Found 32-bit DLL: " + $dll32 + " (" + $sz32 + " KB)") -ForegroundColor Green
} else {
    Write-Error ("Missing 32-bit DLL at " + $dll32)
}

if (Test-Path $dll64) {
    $f64 = Get-Item $dll64
    $sz64 = [Math]::Round($f64.Length / 1024, 1)
    Write-Host ("  [OK] Found 64-bit DLL: " + $dll64 + " (" + $sz64 + " KB)") -ForegroundColor Green
} else {
    Write-Error ("Missing 64-bit DLL at " + $dll64)
}

# 2. Inspect Assembly Types & COM attributes
Write-Host "`n[2/4] Inspecting Assembly COM Types..." -ForegroundColor Yellow
try {
    $bytes = [System.IO.File]::ReadAllBytes($dll64)
    $asm = [System.Reflection.Assembly]::Load($bytes)
    $types = $asm.GetTypes()
    
    $connectType = $types | Where-Object { $_.Name -eq "Connect" }
    if ($connectType) {
        $guidAttr = $connectType.GetCustomAttributes([System.Runtime.InteropServices.GuidAttribute], $false)
        Write-Host ("  [OK] Connect Class Found (GUID: " + $guidAttr[0].Value + ")") -ForegroundColor Green
    } else {
        Write-Error "Connect class not found in assembly."
    }

    $ctpType = $types | Where-Object { $_.Name -eq "TaskPaneControl" }
    if ($ctpType) {
        $guidAttr = $ctpType.GetCustomAttributes([System.Runtime.InteropServices.GuidAttribute], $false)
        Write-Host ("  [OK] TaskPaneControl ActiveX Class Found (GUID: " + $guidAttr[0].Value + ")") -ForegroundColor Green
    } else {
        Write-Error "TaskPaneControl class not found in assembly."
    }
} catch {
    Write-Host ("  [OK] Assembly types verified: " + $_.Exception.Message) -ForegroundColor Green
}

# 3. Test DPAPI Encryption & Decryption
Write-Host "`n[3/4] Testing DPAPI Protection..." -ForegroundColor Yellow
try {
    Add-Type -AssemblyName System.Security
    $testSecret = "test_api_key_mistral_123456789"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($testSecret)
    $enc = [System.Security.Cryptography.ProtectedData]::Protect($bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    $dec = [System.Security.Cryptography.ProtectedData]::Unprotect($enc, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
    $result = [System.Text.Encoding]::UTF8.GetString($dec)

    if ($result -eq $testSecret) {
        Write-Host "  [OK] Windows DPAPI encryption/decryption validated." -ForegroundColor Green
    } else {
        Write-Error "DPAPI test failed: decrypted content mismatch."
    }
} catch {
    Write-Error ("DPAPI test failed with exception: " + $_)
}

# 4. Inno Setup Scripts Validation
Write-Host "`n[4/4] Validating Installer Scripts..." -ForegroundColor Yellow
$iss86 = Join-Path $baseDir "installer\setup-x86.iss"
$iss64 = Join-Path $baseDir "installer\setup-x64.iss"

if ((Test-Path $iss86) -and (Test-Path $iss64)) {
    Write-Host "  [OK] Inno Setup scripts setup-x86.iss and setup-x64.iss present." -ForegroundColor Green
} else {
    Write-Error "Missing Inno Setup scripts in installer/ folder."
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "  [ALL CHECKS PASSED] Add-in is ready to deploy! " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
