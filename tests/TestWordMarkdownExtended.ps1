Add-Type -Path "packages/System.Runtime.CompilerServices.Unsafe.4.5.3/lib/net461/System.Runtime.CompilerServices.Unsafe.dll"
Add-Type -Path "packages/System.Buffers.4.5.1/lib/net461/System.Buffers.dll"
Add-Type -Path "packages/System.Numerics.Vectors.4.5.0/lib/net46/System.Numerics.Vectors.dll"
Add-Type -Path "packages/System.Memory.4.5.4/lib/net461/System.Memory.dll"
Add-Type -Path "packages/Markdig.0.22.0/lib/net452/Markdig.dll"
Add-Type -Path "packages/NetOfficeFw.Core.1.9.10/lib/net462/NetOffice.dll"
Add-Type -Path "packages/NetOfficeFw.Office.1.9.10/lib/net462/OfficeApi.dll"
Add-Type -Path "packages/NetOfficeFw.Word.1.9.10/lib/net462/WordApi.dll"
Add-Type -Path "bin/x64/Release/MistralOfficeAddin.dll"

Write-Host "=========================================================="
Write-Host " Running Extended Word Markdown Renderer Tests"
Write-Host "=========================================================="

$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

try {
    $doc = $wordApp.Documents.Add()
    $sel = $wordApp.Selection
    
    # 1. Test Selection Replacement
    $sel.Font.Name = "Georgia"
    $sel.Font.Size = 13.0
    $sel.TypeText("Original text to be replaced")
    
    # Select the paragraph
    $doc.Paragraphs.Item(1).Range.Select()
    
    $wordCtrl = New-Object MistralOfficeAddin.Hosts.WordController($wordApp)
    $replaceMd = "## Replaced Heading`n`nNew paragraph with **bold replacement**."
    $wordCtrl.ReplaceSelection($replaceMd)
    
    Write-Host "[1/2] Selection replacement verified."
    $content = $doc.Content.Text
    if ($content.Contains("Original text to be replaced")) {
        throw "Assertion failed: Original text was not replaced!"
    }
    
    # 2. Test Complex Formatting
    $complexMd = @"
### Deeply Nested Formatting

This paragraph has ***bold italic*** and **bold with `code` inside** and *italic with **nested bold*** text.

- Level 1 Bullet
  - Level 2 Sub-bullet with `Consolas Code`
  - Level 2 Sub-bullet with **Bold**
- Level 1 Second Bullet

1. Ordered Step 1: Initialize
2. Ordered Step 2: Validate
   1. Sub-step A
   2. Sub-step B
3. Ordered Step 3: Deploy
"@
    $wordCtrl.InsertTextAtCursor($complexMd)
    Write-Host "[2/2] Complex nested formatting inserted without error."
    
    Write-Host ""
    Write-Host "=========================================================="
    Write-Host " [SUCCESS] Extended Word Markdown Tests Passed!"
    Write-Host "=========================================================="
}
finally {
    if ($doc -ne $null) {
        $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    }
    if ($wordApp -ne $null) {
        $wordApp.Quit()
    }
}
