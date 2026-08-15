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
Write-Host " Running Word Markdown Renderer Integration Tests"
Write-Host "=========================================================="

$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

try {
    $doc = $wordApp.Documents.Add()
    $sel = $wordApp.Selection
    
    # Set document font to Arial 12pt to test font inheritance
    $sel.Font.Name = "Arial"
    $sel.Font.Size = 12.0
    
    $netApp = New-Object NetOffice.WordApi.Application($null, $wordApp)
    $wordCtrl = New-Object MistralOfficeAddin.Hosts.WordController($wordApp)
    
    $testMarkdown = @"
# Strategic Overview

This is an introductory paragraph with **bold text**, *italic text*, and ***bold italic text*** along with `inline_func()` and [Example Link](https://example.com).

## Key Initiatives
- First high-priority item with **bold** note
- Second initiative
  - Sub-initiative item
- Third initiative

## Action Items
1. Complete system audit
2. Implement markdown renderer
3. Validate format alignment

## Performance Data

| Metric | Baseline | Target | Status |
| :--- | :---: | :---: | ---: |
| Latency | 250ms | **120ms** | *Exceeded* |
| Availability | 99.9% | 99.99% | *Met* |

> Quality is not an act, it is a habit.

```csharp
public class DataProcessor
{
    public void Process() => Console.WriteLine("Processed");
}
```

---
"@

    Write-Host "[1/4] Calling WordController.InsertTextAtCursor with Markdown..."
    $wordCtrl.InsertTextAtCursor($testMarkdown)
    
    Write-Host "[2/4] Verifying Document Structure..."
    $paraCount = $doc.Paragraphs.Count
    $tableCount = $doc.Tables.Count
    Write-Host "  -> Total Paragraphs: $paraCount"
    Write-Host "  -> Total Tables: $tableCount"
    
    if ($tableCount -ne 1) {
        throw "Assertion failed: Expected 1 table, found $tableCount"
    }
    
    $tbl = $doc.Tables.Item(1)
    Write-Host "  -> Table Rows: $($tbl.Rows.Count), Columns: $($tbl.Columns.Count)"
    if ($tbl.Rows.Count -ne 3 -or $tbl.Columns.Count -ne 4) {
        throw "Assertion failed: Expected 3 rows and 4 columns in table"
    }
    
    Write-Host "[3/4] Verifying Table Header Formatting..."
    $hdrText = $tbl.Cell(1, 1).Range.Text.TrimEnd([char]13, [char]7)
    Write-Host "  -> Cell(1,1) Text: '$hdrText'"
    Write-Host "  -> Cell(1,1) Bold: $($tbl.Cell(1,1).Range.Font.Bold)"
    Write-Host "  -> Cell(1,1) Font: $($tbl.Cell(1,1).Range.Font.Name)"
    if ($tbl.Cell(1,1).Range.Font.Name -ne "Arial") {
        throw "Assertion failed: Expected table font 'Arial', got '$($tbl.Cell(1,1).Range.Font.Name)'"
    }
    
    Write-Host "[4/4] Verifying Clean Text (No raw markdown syntax markers)..."
    $docText = $doc.Content.Text
    if ($docText.Contains("# Strategic Overview") -or $docText.Contains("**bold text**") -or $docText.Contains("| :--- |")) {
        throw "Assertion failed: Raw markdown symbols found in document text!"
    }
    
    Write-Host ""
    Write-Host "=========================================================="
    Write-Host " [SUCCESS] All Word Markdown Renderer Tests Passed!"
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
