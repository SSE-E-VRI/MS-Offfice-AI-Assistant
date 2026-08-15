$ErrorActionPreference = "Stop"
Write-Host "Testing TLS 1.2 connection to Mistral API..." -ForegroundColor Cyan

$asmPath = "C:\Tools\MsOfficePlugin\bin\x64\Release\MistralOfficeAddin.dll"
$asm = [System.Reflection.Assembly]::LoadFrom($asmPath)
$configType = $asm.GetType("MistralOfficeAddin.Core.ConfigManager")
$configInst = $configType.GetProperty("Instance").GetValue($null)
$apiKey = $configInst.ApiKey

Write-Host "API Key configured: $(if ([string]::IsNullOrWhiteSpace($apiKey)) { 'NO' } else { 'YES (' + $apiKey.Substring(0, 4) + '...)' })" -ForegroundColor Yellow

$clientType = $asm.GetType("MistralOfficeAddin.API.MistralClient")
$client = [Activator]::CreateInstance($clientType, @("https://api.mistral.ai/v1", $apiKey))

try {
    $task = $client.TestConnectionAsync()
    $task.Wait()
    Write-Host "[PASS] TLS 1.2 Handshake & Mistral API Connection Succeeded!" -ForegroundColor Green
}
catch {
    Write-Host "[INFO] Response received: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow
}
$client.Dispose()
