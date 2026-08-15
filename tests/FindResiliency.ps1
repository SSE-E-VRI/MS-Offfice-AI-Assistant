$ErrorActionPreference = "SilentlyContinue"

Write-Host "--- Scanning HKCU:\Software\Microsoft\Office for Resiliency ---"
Get-ChildItem -Path "HKCU:\Software\Microsoft\Office" -Recurse | Where-Object { $_.Name -like "*Resiliency*" } | ForEach-Object {
    Write-Host "Found Resiliency Key: $($_.Name)"
    $key = $_
    $props = Get-ItemProperty -Path $key.PSPath
    $props.PSObject.Properties | ForEach-Object {
        if ($_.Name -notlike "PS*") {
            Write-Host "   Property: $($_.Name)"
        }
    }
}

Write-Host "--- Scanning Word Addins in HKCU and HKLM ---"
Get-ItemProperty "HKCU:\Software\Microsoft\Office\Word\Addins\*" | Select-Object PSChildName, FriendlyName, LoadBehavior | Format-Table -AutoSize
Get-ItemProperty "HKLM:\Software\Microsoft\Office\Word\Addins\*" | Select-Object PSChildName, FriendlyName, LoadBehavior | Format-Table -AutoSize
Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Office\Word\Addins\*" | Select-Object PSChildName, FriendlyName, LoadBehavior | Format-Table -AutoSize
