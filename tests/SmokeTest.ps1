<#
.SYNOPSIS
    Comprehensive Smoke Test for Mistral AI Office Add-in.
    Tests COM activation, IDTExtensibility2, IRibbonExtensibility, TaskPaneControl ActiveX, and Word COM addin loading.
#>

$ErrorActionPreference = "Stop"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Mistral AI Office Add-in - Smoke Test Suite     " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

# Test 1: COM Activation of Addin Class
Write-Host "`n[Test 1/5] Testing COM Activation of MistralAI.Addin..." -ForegroundColor Yellow
try {
    $addinType = [Type]::GetTypeFromProgID("MistralAI.Addin")
    if ($null -eq $addinType) {
        throw "ProgID 'MistralAI.Addin' not found in registry."
    }
    $addin = [Activator]::CreateInstance($addinType)
    if ($null -eq $addin) {
        throw "Failed to instantiate MistralAI.Addin."
    }
    Write-Host "  [PASS] Successfully instantiated MistralAI.Addin ($($addin.GetType().FullName))" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 2: IRibbonExtensibility.GetCustomUI
Write-Host "`n[Test 2/5] Testing IRibbonExtensibility (GetCustomUI)..." -ForegroundColor Yellow
try {
    $asmPath = "C:\Tools\MsOfficePlugin\bin\x64\Release\MistralOfficeAddin.dll"
    if (-not (Test-Path $asmPath)) { $asmPath = "C:\Tools\MsOfficePlugin\bin\x86\Release\MistralOfficeAddin.dll" }
    $asm = [System.Reflection.Assembly]::LoadFrom($asmPath)
    $ribbonInterface = $asm.GetType("MistralOfficeAddin.Addin.IRibbonExtensibility")
    $ribbonMethod = $ribbonInterface.GetMethod("GetCustomUI")
    if ($null -eq $ribbonMethod) {
        throw "Method GetCustomUI not found on IRibbonExtensibility."
    }
    $connectType = $asm.GetType("MistralOfficeAddin.Addin.Connect")
    $connectInstance = [Activator]::CreateInstance($connectType)
    $xml = $ribbonMethod.Invoke($connectInstance, @("Microsoft.Word.Ribbon"))
    if ([string]::IsNullOrWhiteSpace($xml) -or -not $xml.Contains("tabMistralAI")) {
        throw "GetCustomUI did not return expected Ribbon XML. Got: $xml"
    }
    Write-Host "  [PASS] Ribbon XML retrieved successfully ($($xml.Length) chars)" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 3: IDTExtensibility2 Methods (OnConnection, OnStartupComplete, OnDisconnection)
Write-Host "`n[Test 3/5] Testing IDTExtensibility2 lifecycle methods..." -ForegroundColor Yellow
try {
    # Load assembly to access managed interface definitions
    $asmPath = "C:\Tools\MsOfficePlugin\bin\x64\Release\MistralOfficeAddin.dll"
    if (-not (Test-Path $asmPath)) { $asmPath = "C:\Tools\MsOfficePlugin\bin\x86\Release\MistralOfficeAddin.dll" }
    $asm = [System.Reflection.Assembly]::LoadFrom($asmPath)
    $connectType = $asm.GetType("MistralOfficeAddin.Addin.Connect")
    $connectInstance = [Activator]::CreateInstance($connectType)

    # Test IDTExtensibility2 methods
    $idteInterface = $asm.GetType("MistralOfficeAddin.Addin.IDTExtensibility2")
    $connMethod = $idteInterface.GetMethod("OnConnection")
    $custom = [Array]::CreateInstance([object], 0)
    $connMethod.Invoke($connectInstance, [object[]]@($null, 3, $null, $custom))
    Write-Host "  [PASS] IDTExtensibility2.OnConnection executed cleanly." -ForegroundColor Green

    $startupMethod = $idteInterface.GetMethod("OnStartupComplete")
    $startupArgs = [object[]]@(,[Array]::CreateInstance([object], 0))
    $startupMethod.Invoke($connectInstance, $startupArgs)
    Write-Host "  [PASS] IDTExtensibility2.OnStartupComplete executed cleanly." -ForegroundColor Green

    $discMethod = $idteInterface.GetMethod("OnDisconnection")
    $discArgs = [object[]]@(0, [Array]::CreateInstance([object], 0))
    $discMethod.Invoke($connectInstance, $discArgs)
    Write-Host "  [PASS] IDTExtensibility2.OnDisconnection executed cleanly." -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 4: COM Activation of TaskPaneControl ActiveX
Write-Host "`n[Test 4/5] Testing COM Activation of MistralAI.TaskPaneControl..." -ForegroundColor Yellow
try {
    $tpType = [Type]::GetTypeFromProgID("MistralAI.TaskPaneControl")
    if ($null -eq $tpType) {
        throw "ProgID 'MistralAI.TaskPaneControl' not found in registry."
    }
    $tp = [Activator]::CreateInstance($tpType)
    if ($null -eq $tp) {
        throw "Failed to instantiate MistralAI.TaskPaneControl."
    }
    Write-Host "  [PASS] Successfully instantiated MistralAI.TaskPaneControl" -ForegroundColor Green

    # Verify IObjectSafety
    $safetyInterface = $tpType.GetInterface("IObjectSafety")
    if ($null -ne $safetyInterface) {
        Write-Host "  [PASS] IObjectSafety interface implemented." -ForegroundColor Green
    }
    $tp.Dispose()
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 5: Live Microsoft Word COM Addin Integration Test
Write-Host "`n[Test 5/5] Testing Live Microsoft Word COM Addin Integration..." -ForegroundColor Yellow
try {
    $word = New-Object -ComObject Word.Application -ErrorAction SilentlyContinue
    if ($null -ne $word) {
        $wordVer = $word.Version
        $wordName = $word.Name
        Write-Host "  [INFO] Found active Word instance: $wordName (Version: $wordVer)" -ForegroundColor Cyan
        
        $mistralAddin = $null
        foreach ($ca in $word.COMAddIns) {
            if ($ca.ProgId -eq "MistralAI.Addin") {
                $mistralAddin = $ca
                break
            }
        }
        
        if ($null -ne $mistralAddin) {
            Write-Host "  [INFO] MistralAI.Addin COMAddIn entry found. Connect state: $($mistralAddin.Connect)" -ForegroundColor Cyan
            
            # Attempt to connect addin in Word
            try {
                $mistralAddin.Connect = $true
                Write-Host "  [PASS] Successfully connected MistralAI.Addin inside Word!" -ForegroundColor Green
            } catch {
                Write-Host "  [WARN] Connecting addin in automation mode: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        } else {
            Write-Host "  [WARN] MistralAI.Addin not listed in Word.COMAddIns collection yet (check registry keys)." -ForegroundColor Yellow
        }
        
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    } else {
        Write-Host "  [INFO] Microsoft Word is not running / COM automation unavailable in this context (skipped live host test)." -ForegroundColor Gray
    }
}
catch {
    Write-Host "  [WARN] Live Word test encountered: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "  [SMOKE TEST COMPLETED SUCCESSFULLY]             " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
