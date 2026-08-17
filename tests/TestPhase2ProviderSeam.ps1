# Verification test for Phase 2 Provider Seam & Orchestrator
$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Testing Phase 2 Provider Contract & ChatOrchestrator     " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$dllPath = Join-Path $PSScriptRoot "..\bin\x64\Release\MistralOfficeAddin.dll"
if (-not (Test-Path $dllPath)) {
    $dllPath = Join-Path $PSScriptRoot "..\bin\x86\Release\MistralOfficeAddin.dll"
}

Write-Host "Loading assembly: $dllPath" -ForegroundColor Yellow
$bytes = [System.IO.File]::ReadAllBytes($dllPath)
$asm = [System.Reflection.Assembly]::Load($bytes)

$passed = 0
$failed = 0

function Assert-Condition($name, [bool]$condition) {
    if ($condition) {
        Write-Host "  [PASS] $name" -ForegroundColor Green
        $global:passed++
    } else {
        Write-Host "  [FAIL] $name" -ForegroundColor Red
        $global:failed++
    }
}

# 1. Verify Types exist
$tIAIProvider = $asm.GetType("MistralOfficeAddin.Providers.IAIProvider")
Assert-Condition "IAIProvider interface exists" ($tIAIProvider -ne $null)

$tMistralProvider = $asm.GetType("MistralOfficeAddin.Providers.MistralProvider")
Assert-Condition "MistralProvider class exists" ($tMistralProvider -ne $null)

$tChatOrchestrator = $asm.GetType("MistralOfficeAddin.Providers.ChatOrchestrator")
Assert-Condition "ChatOrchestrator class exists" ($tChatOrchestrator -ne $null)

$tProviderFactory = $asm.GetType("MistralOfficeAddin.Providers.ProviderFactory")
Assert-Condition "ProviderFactory class exists" ($tProviderFactory -ne $null)

# 2. Test MistralProvider instantiation & default model list
$mistral = [Activator]::CreateInstance($tMistralProvider, @("https://api.mistral.ai/v1", "dummy_key"))
Assert-Condition "MistralProvider instantiated" ($mistral -ne $null)
Assert-Condition "MistralProvider.ProviderType is Mistral" ($mistral.ProviderType.ToString() -eq "Mistral")

$isVisionPixtral = $mistral.CheckVisionSupport("pixtral-large-latest")
Assert-Condition "CheckVisionSupport('pixtral-large-latest') == true" ($isVisionPixtral -eq $true)

$isVisionSmall = $mistral.CheckVisionSupport("mistral-small-latest")
Assert-Condition "CheckVisionSupport('mistral-small-latest') == false" ($isVisionSmall -eq $false)

# 3. Test ChatOrchestrator
$orchestrator = [Activator]::CreateInstance($tChatOrchestrator, @($mistral))
Assert-Condition "ChatOrchestrator instantiated with MistralProvider" ($orchestrator.CurrentProvider -ne $null)
Assert-Condition "ChatOrchestrator.IsStreaming is false initially" ($orchestrator.IsStreaming -eq $false)

# 4. Test safe provider replacement
$newMistral = [Activator]::CreateInstance($tMistralProvider, @("https://api.mistral.ai/v1", "new_key"))
$orchestrator.UpdateProvider($newMistral)
Assert-Condition "ChatOrchestrator.UpdateProvider replaced provider" ($orchestrator.CurrentProvider -eq $newMistral)

# 5. Clean up
$orchestrator.Dispose()
Assert-Condition "ChatOrchestrator disposed safely" $true

# 6. Verify Assembly metadata (Version 0.0.0.0, Product "AI Assistant")
$titleAttr = $asm.GetCustomAttributes([System.Reflection.AssemblyTitleAttribute], $false)
Assert-Condition "AssemblyTitle is 'AI Assistant'" ($titleAttr.Length -gt 0 -and $titleAttr[0].Title -eq "AI Assistant")

$ver = $asm.GetName().Version.ToString()
Assert-Condition "AssemblyVersion is 0.0.0.0" ($ver -eq "0.0.0.0")

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Results: $passed Passed, $failed Failed" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
Write-Host "============================================================" -ForegroundColor Cyan

if ($failed -gt 0) {
    exit 1
}
