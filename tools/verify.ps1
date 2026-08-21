# Verifies that the Mistral AI COM add-in loads and connects in
# Word, Excel and PowerPoint via COM automation.
# Exits with the number of failed apps (0 = all good).

$ErrorActionPreference = 'Continue'
$failures = 0

$apps = @(
    @{ ProgId = 'Word.Application';     Name = 'Word' },
    @{ ProgId = 'Excel.Application';    Name = 'Excel' },
    @{ ProgId = 'PowerPoint.Application'; Name = 'PowerPoint' }
)

foreach ($a in $apps) {
    $app = $null
    try {
        $app = New-Object -ComObject $a.ProgId
    } catch {
        Write-Output ("FAIL {0}: cannot start host ({1})" -f $a.Name, $_.Exception.Message)
        $failures++
        continue
    }

    $found = $false
    $connected = $false
    $desc = ''
    try {
        foreach ($ai in $app.COMAddIns) {
            if ($ai.ProgId -eq 'MSOfficeAIAssistant.Addin') {
                $found = $true
                $connected = $ai.Connect
                $desc = $ai.Description
                break
            }
        }
    } catch { }

    if ($found -and $connected) {
        Write-Output ("PASS {0}: MSOfficeAIAssistant.Addin is loaded and connected ({1})" -f $a.Name, $desc)
    } elseif ($found) {
        Write-Output ("WARN {0}: MSOfficeAIAssistant.Addin is registered but NOT connected - check LoadBehavior" -f $a.Name)
        $failures++
    } else {
        Write-Output ("FAIL {0}: MSOfficeAIAssistant.Addin not present in COMAddIns - run install.cmd" -f $a.Name)
        $failures++
    }

    try { $app.Quit() } catch { }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()
exit $failures
