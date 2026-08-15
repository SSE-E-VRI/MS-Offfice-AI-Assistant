$bytes = [System.IO.File]::ReadAllBytes("c:\Tools\MsOfficePlugin\bin\x64\Release\MistralOfficeAddin.dll")
$asm = [System.Reflection.Assembly]::Load($bytes)
$asm.GetManifestResourceNames()
