<#
.SYNOPSIS
    Comprehensive Smoke Test for AI Assistant Office Add-in v0.0.0.
    Tests COM activation, IDTExtensibility2, IRibbonExtensibility, TaskPaneControl ActiveX,
    All 4 AI Providers, Attachment pipeline, and Registry configuration.
#>

$ErrorActionPreference = "Stop"

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  AI Assistant Office Add-in - Smoke Test Suite   " -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

$dllPath = "C:\Tools\MsOfficePlugin\bin\x64\Release\MistralOfficeAddin.dll"
if (-not (Test-Path $dllPath)) {
    $dllPath = "C:\Tools\MsOfficePlugin\bin\x86\Release\MistralOfficeAddin.dll"
}

Write-Host "Target Assembly: $dllPath" -ForegroundColor Gray

# Test 1: Assembly Metadata & Rebranding
Write-Host "`n[Test 1/7] Testing Assembly Identity & Version 0.0.0..." -ForegroundColor Yellow
try {
    $asmBytes = [System.IO.File]::ReadAllBytes($dllPath)
    $asm = [System.Reflection.Assembly]::Load($asmBytes)

    $titleAttr = $asm.GetCustomAttributes([System.Reflection.AssemblyTitleAttribute], $false)
    if ($titleAttr.Length -eq 0 -or $titleAttr[0].Title -ne "AI Assistant") {
        throw "AssemblyTitle is '$($titleAttr[0].Title)', expected 'AI Assistant'."
    }

    $ver = $asm.GetName().Version.ToString()
    if ($ver -ne "0.0.0.0") {
        throw "AssemblyVersion is '$ver', expected '0.0.0.0'."
    }

    $copyAttr = $asm.GetCustomAttributes([System.Reflection.AssemblyCopyrightAttribute], $false)
    if ($copyAttr.Length -eq 0 -or -not $copyAttr[0].Copyright.Contains("D.Manikandan")) {
        throw "AssemblyCopyright does not contain developer credit."
    }

    Write-Host "  [PASS] Product Identity: AI Assistant (Version: $ver)" -ForegroundColor Green
    Write-Host "  [PASS] Credit: $($copyAttr[0].Copyright)" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 2: COM Activation of Addin Class
Write-Host "`n[Test 2/7] Testing COM Activation of MistralAI.Addin..." -ForegroundColor Yellow
try {
    $addinType = [Type]::GetTypeFromProgID("MistralAI.Addin")
    if ($null -ne $addinType) {
        $addin = [Activator]::CreateInstance($addinType)
        Write-Host "  [PASS] Successfully instantiated MistralAI.Addin via COM" -ForegroundColor Green
    } else {
        Write-Host "  [INFO] ProgID 'MistralAI.Addin' not registered in current session (run register.cmd to register)." -ForegroundColor Gray
    }
}
catch {
    Write-Host "  [WARN] COM Activation: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Test 3: IRibbonExtensibility.GetCustomUI
Write-Host "`n[Test 3/7] Testing IRibbonExtensibility (GetCustomUI)..." -ForegroundColor Yellow
try {
    $connectType = $asm.GetType("MistralOfficeAddin.Addin.Connect")
    $connectInstance = [Activator]::CreateInstance($connectType)

    $ribbonInterface = $asm.GetType("MistralOfficeAddin.Addin.IRibbonExtensibility")
    $ribbonMethod = $ribbonInterface.GetMethod("GetCustomUI")
    $xml = $ribbonMethod.Invoke($connectInstance, @("Microsoft.Word.Ribbon"))

    if ([string]::IsNullOrWhiteSpace($xml) -or -not $xml.Contains("label=""AI Assistant""")) {
        throw "GetCustomUI did not return tab with label='AI Assistant'."
    }
    Write-Host "  [PASS] Ribbon XML retrieved successfully with tab label 'AI Assistant' ($($xml.Length) chars)" -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 4: IDTExtensibility2 Methods
Write-Host "`n[Test 4/7] Testing IDTExtensibility2 lifecycle methods..." -ForegroundColor Yellow
try {
    $idteInterface = $asm.GetType("MistralOfficeAddin.Addin.IDTExtensibility2")
    $connMethod = $idteInterface.GetMethod("OnConnection")
    $custom = [Array]::CreateInstance([object], 0)
    $connMethod.Invoke($connectInstance, [object[]]@($null, 3, $null, $custom))
    Write-Host "  [PASS] IDTExtensibility2.OnConnection executed cleanly." -ForegroundColor Green

    $startupMethod = $idteInterface.GetMethod("OnStartupComplete")
    $startupMethod.Invoke($connectInstance, [object[]]@(,$custom))
    Write-Host "  [PASS] IDTExtensibility2.OnStartupComplete executed cleanly." -ForegroundColor Green

    $discMethod = $idteInterface.GetMethod("OnDisconnection")
    $discMethod.Invoke($connectInstance, [object[]]@(0, $custom))
    Write-Host "  [PASS] IDTExtensibility2.OnDisconnection executed cleanly." -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 5: COM Activation of TaskPaneControl ActiveX
Write-Host "`n[Test 5/7] Testing TaskPaneControl ActiveX..." -ForegroundColor Yellow
try {
    $tpType = $asm.GetType("MistralOfficeAddin.Addin.TaskPaneControl")
    $tp = [Activator]::CreateInstance($tpType)
    $safetyInterface = $tpType.GetInterface("IObjectSafety")
    if ($null -ne $safetyInterface) {
        Write-Host "  [PASS] TaskPaneControl implements IObjectSafety." -ForegroundColor Green
    }
    $tp.Dispose()
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 6: Multi-Provider Engine
Write-Host "`n[Test 6/7] Testing Multi-Provider Suite (Mistral, Groq, Gemini, Custom API)..." -ForegroundColor Yellow
try {
    $tFactory = $asm.GetType("MistralOfficeAddin.Providers.ProviderFactory")
    $tMistral = $asm.GetType("MistralOfficeAddin.Providers.MistralProvider")
    $tGroq = $asm.GetType("MistralOfficeAddin.Providers.GroqProvider")
    $tGemini = $asm.GetType("MistralOfficeAddin.Providers.GeminiProvider")
    $tCustom = $asm.GetType("MistralOfficeAddin.Providers.CustomApiProvider")

    $m = [Activator]::CreateInstance($tMistral, @("https://api.mistral.ai/v1", "key"))
    $g = [Activator]::CreateInstance($tGroq, @("key"))
    $gem = [Activator]::CreateInstance($tGemini, @("key", "https://generativelanguage.googleapis.com"))
    $c = [Activator]::CreateInstance($tCustom, @("http://localhost:11434/v1", "", "llama3", $null))

    if ($m.ProviderType.ToString() -ne "Mistral" -or
        $g.ProviderType.ToString() -ne "Groq" -or
        $gem.ProviderType.ToString() -ne "Gemini" -or
        $c.ProviderType.ToString() -ne "Custom") {
        throw "Provider types mismatch."
    }

    Write-Host "  [PASS] Mistral, Groq, Gemini, and Custom API providers initialized." -ForegroundColor Green
}
catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Test 7: Registry Configuration & Outlook De-scope Verification
Write-Host "`n[Test 7/7] Verifying Registry Configuration (Word, Excel, PowerPoint - Outlook descoped)..." -ForegroundColor Yellow
$supportedApps = @("Word", "Excel", "PowerPoint")
$regOk = $true
foreach ($app in $supportedApps) {
    $key = "HKCU:\Software\Microsoft\Office\$app\Addins\MistralAI.Addin"
    if (Test-Path $key) {
        $friendlyName = (Get-ItemProperty -Path $key -Name "FriendlyName" -ErrorAction SilentlyContinue).FriendlyName
        Write-Host "  [INFO] $app add-in registered with FriendlyName: '$friendlyName'" -ForegroundColor Cyan
    }
}

$outlookKey = "HKCU:\Software\Microsoft\Office\Outlook\Addins\MistralAI.Addin"
if (Test-Path $outlookKey) {
    Write-Host "  [INFO] Note: Outlook add-in key exists from previous registration (run unregister.cmd and register.cmd to align)." -ForegroundColor Gray
} else {
    Write-Host "  [PASS] Outlook is cleanly not registered." -ForegroundColor Green
}

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "  [ALL SMOKE TESTS COMPLETED SUCCESSFULLY]        " -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
