$ErrorActionPreference = "SilentlyContinue"
$asm = [System.Reflection.Assembly]::LoadFrom("C:\Tools\MsOfficePlugin\packages\NetOfficeFw.Office.1.9.10\lib\net462\OfficeApi.dll")
foreach ($t in $asm.GetExportedTypes()) {
    if ($t.Name -like "*CTP*" -or $t.Name -like "*TaskPane*" -or $t.Name -like "*Consumer*") {
        Write-Host "Found: $($t.FullName)"
        Write-Host "  IsInterface: $($t.IsInterface)"
        foreach ($m in $t.GetMethods()) {
            Write-Host "  Method: $($m.Name)"
        }
    }
}
