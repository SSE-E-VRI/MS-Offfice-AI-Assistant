$ErrorActionPreference = "SilentlyContinue"
$asm = [System.Reflection.Assembly]::LoadFrom("C:\Tools\MsOfficePlugin\packages\NetOfficeFw.Office.1.9.10\lib\net462\OfficeApi.dll")
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -like "*CTP*" -or $t.Name -like "*TaskPane*" -or $t.Name -like "*CustomTaskPane*") {
        Write-Host "Type: $($t.FullName)"
        Write-Host "  IsInterface: $($t.IsInterface)"
        foreach ($attr in $t.GetCustomAttributes($false)) {
            Write-Host "  Attr: $($attr.GetType().FullName)"
        }
        foreach ($m in $t.GetMethods()) {
            Write-Host "  Method: $($m.Name)"
        }
    }
}
