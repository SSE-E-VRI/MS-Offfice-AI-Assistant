$ErrorActionPreference = "SilentlyContinue"
$asm = [System.Reflection.Assembly]::LoadFrom("C:\Tools\MsOfficePlugin\packages\NetOfficeFw.Office.1.9.10\lib\net462\OfficeApi.dll")
Write-Host "Total types: $($asm.GetTypes().Length)"
$asm.GetTypes() | Where-Object { $_.IsInterface } | ForEach-Object {
    Write-Host "Interface: $($_.FullName)"
}
