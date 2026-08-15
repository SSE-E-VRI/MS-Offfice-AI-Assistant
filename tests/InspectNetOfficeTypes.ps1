$ErrorActionPreference = "SilentlyContinue"
$asmPath = "C:\Tools\MsOfficePlugin\packages\NetOfficeFw.Office.1.9.10\lib\net462\OfficeApi.dll"
$asm = [System.Reflection.Assembly]::LoadFrom($asmPath)
$types = $asm.GetExportedTypes()
$types | Where-Object { $_.Name -like "*Extensib*" -or $_.Name -like "*IDT*" -or $_.Name -like "*Ribbon*" -or $_.Name -like "*TaskPane*" } | ForEach-Object {
    Write-Host "Type: $($_.FullName)"
    if ($_.IsInterface) {
        Write-Host "  Interface Guid: $([System.Runtime.InteropServices.Marshal]::GenerateGuidForType($_))"
        $att = [System.Attribute]::GetCustomAttributes($_) | ForEach-Object { Write-Host "  Attr: $($_.GetType().Name)" }
    }
}
