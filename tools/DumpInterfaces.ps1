# Dumps interface GUIDs + member order from the Office Core PIA (Office.dll)
# Used to verify our hand-declared ComImport interfaces match the real type library.
$ErrorActionPreference = 'Stop'
$asm = [Reflection.Assembly]::LoadFrom('C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\office.dll')

$wanted = @('ICTPFactory','ICustomTaskPane','ICustomTaskPaneConsumer','IRibbonExtensibility','IRibbonControl')

foreach ($name in $wanted) {
    $t = $asm.GetTypes() | Where-Object { $_.Name -eq $name -and $_.IsInterface }
    foreach ($iface in $t) {
        Write-Output ('=== ' + $iface.FullName + ' ===')
        Write-Output ('Guid: ' + $iface.GUID.ToString('B').ToUpper())
        # InterfaceType: check ComInterfaceType from ComImportAttribute is not directly readable;
        # member order from GetMethods reflects vtable order for the PIA.
        $methods = $iface.GetMethods() | Sort-Object { [int]($_.MetadataToken) }
        foreach ($m in $methods) {
            $params = ($m.GetParameters() | ForEach-Object { $_.ParameterType.Name + ' ' + $_.Name }) -join ', '
            Write-Output ('  ' + $m.ReturnType.Name + ' ' + $m.Name + '(' + $params + ')')
        }
        Write-Output ''
    }
}
