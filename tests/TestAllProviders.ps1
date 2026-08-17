# Verification test for All AI Providers (Mistral, Groq, Gemini, Custom API)
$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Testing All Providers & Multi-Provider Config Engine      " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$dllPath = Join-Path $PSScriptRoot "..\bin\x64\Release\MistralOfficeAddin.dll"
if (-not (Test-Path $dllPath)) {
    $dllPath = Join-Path $PSScriptRoot "..\bin\x86\Release\MistralOfficeAddin.dll"
}

$dllDir = Split-Path $dllPath
$jsonPkg = Join-Path $PSScriptRoot "..\packages\Newtonsoft.Json.13.0.4\lib\net45\Newtonsoft.Json.dll"
if (Test-Path $jsonPkg) {
    [System.Reflection.Assembly]::LoadFrom($jsonPkg) | Out-Null
}

$bytes = [System.IO.File]::ReadAllBytes($dllPath)
$asm = [System.Reflection.Assembly]::Load($bytes)

$global:passed = 0
$global:failed = 0

function Assert-Condition($name, [bool]$condition) {
    if ($condition) {
        Write-Host "  [PASS] $name" -ForegroundColor Green
        $global:passed++
    } else {
        Write-Host "  [FAIL] $name" -ForegroundColor Red
        $global:failed++
    }
}

# 1. Verify Provider Types
$tMistral = $asm.GetType("MistralOfficeAddin.Providers.MistralProvider")
$tGroq = $asm.GetType("MistralOfficeAddin.Providers.GroqProvider")
$tGemini = $asm.GetType("MistralOfficeAddin.Providers.GeminiProvider")
$tCustom = $asm.GetType("MistralOfficeAddin.Providers.CustomApiProvider")
$tFactory = $asm.GetType("MistralOfficeAddin.Providers.ProviderFactory")
$tClient = $asm.GetType("MistralOfficeAddin.Providers.OpenAICompatibleClient")
$tGeminiParser = $asm.GetType("MistralOfficeAddin.Providers.GeminiStreamingParser")

Assert-Condition "MistralProvider exists" ($tMistral -ne $null)
Assert-Condition "GroqProvider exists" ($tGroq -ne $null)
Assert-Condition "GeminiProvider exists" ($tGemini -ne $null)
Assert-Condition "CustomApiProvider exists" ($tCustom -ne $null)

# 2. Test Groq Provider Instantiation & Models
$groq = [Activator]::CreateInstance($tGroq, @("gsk_dummy"))
Assert-Condition "GroqProvider instantiated" ($groq -ne $null)
Assert-Condition "Groq.ProviderType is Groq" ($groq.ProviderType.ToString() -eq "Groq")
Assert-Condition "Groq supports vision on llama-3.2-11b-vision-preview" ($groq.CheckVisionSupport("llama-3.2-11b-vision-preview") -eq $true)

# 3. Test Gemini Provider Instantiation & Models
$gemini = [Activator]::CreateInstance($tGemini, @("gemini_key", "https://generativelanguage.googleapis.com"))
Assert-Condition "GeminiProvider instantiated" ($gemini -ne $null)
Assert-Condition "Gemini.ProviderType is Gemini" ($gemini.ProviderType.ToString() -eq "Gemini")
Assert-Condition "Gemini natively supports vision" ($gemini.CheckVisionSupport("gemini-1.5-flash") -eq $true)

# 4. Test Custom API Provider with Localhost
$custom = [Activator]::CreateInstance($tCustom, @("http://127.0.0.1:11434/v1", "", "llama3", $null))
Assert-Condition "CustomApiProvider instantiated on loopback" ($custom -ne $null)
Assert-Condition "Custom.ProviderType is Custom" ($custom.ProviderType.ToString() -eq "Custom")

# 5. Test HTTPS Validation on Custom API
$valMethod = $tClient.GetMethod("ValidateEndpointUrl", [System.Reflection.BindingFlags]"Public,Static")

$httpsOk = $false
try {
    $valMethod.Invoke($null, @("https://api.openai.com/v1"))
    $httpsOk = $true
} catch { $httpsOk = $false }
Assert-Condition "ValidateEndpointUrl accepts HTTPS remote" $httpsOk

$httpLoopbackOk = $false
try {
    $valMethod.Invoke($null, @("http://localhost:11434/v1"))
    $httpLoopbackOk = $true
} catch { $httpLoopbackOk = $false }
Assert-Condition "ValidateEndpointUrl accepts HTTP loopback (localhost)" $httpLoopbackOk

$httpRemoteBlocked = $false
try {
    $valMethod.Invoke($null, @("http://insecure-api.example.com/v1"))
} catch {
    $httpRemoteBlocked = $true
}
Assert-Condition "ValidateEndpointUrl blocks insecure HTTP remote domain" $httpRemoteBlocked

# 6. Test Gemini SSE Streaming Parser
$testLine = 'data: {"candidates": [{"content": {"parts": [{"text": "Hello world!"}], "role": "model"}}]}'
$parseParams = [object[]]@($testLine, [string]::Empty, $false)
$parseMethod = $tGeminiParser.GetMethod("TryParseLine", [System.Reflection.BindingFlags]"Public,Static")
$parseSuccess = $parseMethod.Invoke($null, $parseParams)

Assert-Condition "GeminiStreamingParser parsed sample SSE delta" ($parseSuccess -eq $true -and $parseParams[1] -eq "Hello world!")

# 7. Test ProviderFactory
$tConfig = $asm.GetType("MistralOfficeAddin.Core.ConfigManager")
$configInst = $tConfig.GetProperty("Instance").GetValue($null, $null)
$activeProv = $tFactory.GetMethod("CreateFromConfig", [System.Reflection.BindingFlags]"Public,Static").Invoke($null, @($configInst))
Assert-Condition "ProviderFactory creates active provider instance" ($activeProv -ne $null)

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Results: $global:passed Passed, $global:failed Failed" -ForegroundColor $(if ($global:failed -eq 0) { "Green" } else { "Red" })
Write-Host "============================================================" -ForegroundColor Cyan

if ($global:failed -gt 0) {
    exit 1
}
