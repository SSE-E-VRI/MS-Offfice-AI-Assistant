# Verification test for Attachment Pipeline & Open XML Extractors
$ErrorActionPreference = "Stop"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Testing Attachment Pipeline & Document Extractors         " -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

$dllPath = Join-Path $PSScriptRoot "..\bin\x64\Release\MistralOfficeAddin.dll"
if (-not (Test-Path $dllPath)) {
    $dllPath = Join-Path $PSScriptRoot "..\bin\x86\Release\MistralOfficeAddin.dll"
}

$dllDir = Split-Path $dllPath
$jsonPkg = Join-Path $PSScriptRoot "..\packages\Newtonsoft.Json.13.0.4\lib\net45\Newtonsoft.Json.dll"
if (Test-Path $jsonPkg) { [System.Reflection.Assembly]::LoadFrom($jsonPkg) | Out-Null }

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

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

$tExtractor = $asm.GetType("MistralOfficeAddin.Attachments.AttachmentExtractor")
Assert-Condition "AttachmentExtractor type exists" ($tExtractor -ne $null)

$tempDir = Join-Path $PSScriptRoot "temp_att_test"
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
New-Item -ItemType Directory -Path $tempDir | Out-Null

try {
    # 1. Test Plain Text Extraction
    $txtFile = [string](Join-Path $tempDir "sample.txt")
    [System.IO.File]::WriteAllText($txtFile, "Hello Office AI Assistant attachment text test.")
    
    $extractTask = $tExtractor.GetMethod("ExtractAsync", [System.Reflection.BindingFlags]"Public,Static").Invoke($null, [object[]]@($txtFile))
    $block = $extractTask.GetAwaiter().GetResult()

    Assert-Condition "Text extraction returns correct content" ($block.ExtractedText.Contains("Hello Office AI Assistant"))
    Assert-Condition "Text extraction IsImage is false" ($block.IsImage -eq $false)

    # 2. Test Image Extraction
    $imgFile = [string](Join-Path $tempDir "sample.png")
    [System.IO.File]::WriteAllBytes($imgFile, [byte[]]@(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))
    
    $imgTask = $tExtractor.GetMethod("ExtractAsync", [System.Reflection.BindingFlags]"Public,Static").Invoke($null, [object[]]@($imgFile))
    $imgBlock = $imgTask.GetAwaiter().GetResult()

    Assert-Condition "Image extraction IsImage is true" ($imgBlock.IsImage -eq $true)
    Assert-Condition "Image ContentType is image/png" ($imgBlock.ContentType -eq "image/png")
    Assert-Condition "Image RawBytes length matches" ($imgBlock.RawBytes.Length -eq 8)

    # 3. Test Legacy Rejection (.doc)
    $legacyFile = [string](Join-Path $tempDir "legacy.doc")
    [System.IO.File]::WriteAllText($legacyFile, "legacy binary data")
    
    $legacyBlocked = $false
    try {
        $tExtractor.GetMethod("ExtractAsync", [System.Reflection.BindingFlags]"Public,Static").Invoke($null, [object[]]@($legacyFile)).GetAwaiter().GetResult()
    } catch {
        $legacyBlocked = $true
    }
    Assert-Condition "Legacy .doc file is rejected with exception" $legacyBlocked

    # 4. Test Mock OpenXML Docx
    $docxFile = [string](Join-Path $tempDir "test.docx")
    $docXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>OpenXML Word Extraction Works!</w:t></w:r></w:p>
  </w:body>
</w:document>
"@
    # Create valid ZIP container with word/document.xml
    $zipArchive = [System.IO.Compression.ZipFile]::Open($docxFile, [System.IO.Compression.ZipArchiveMode]::Create)
    $entry = $zipArchive.CreateEntry("word/document.xml")
    $sw = New-Object System.IO.StreamWriter($entry.Open(), [System.Text.Encoding]::UTF8)
    $sw.Write($docXml)
    $sw.Dispose()
    $zipArchive.Dispose()

    $docxTask = $tExtractor.GetMethod("ExtractAsync", [System.Reflection.BindingFlags]"Public,Static").Invoke($null, [object[]]@($docxFile))
    $docxBlock = $docxTask.GetAwaiter().GetResult()
    Assert-Condition "Docx extraction extracts paragraphs" ($docxBlock.ExtractedText.Contains("OpenXML Word Extraction Works!"))

} finally {
    if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Results: $global:passed Passed, $global:failed Failed" -ForegroundColor $(if ($global:failed -eq 0) { "Green" } else { "Red" })
Write-Host "============================================================" -ForegroundColor Cyan

if ($global:failed -gt 0) {
    exit 1
}
