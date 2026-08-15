; Inno Setup 6 Script for Mistral AI Office Add-in (32-bit / x86)
; Target: 32-bit Microsoft Office (Word, Excel, PowerPoint, Outlook)

#define MyAppName "Mistral AI Office Add-in (32-bit)"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Mistral AI Community"
#define MyAppURL "https://mistral.ai"
#define MyAppExeName "MistralOfficeAddin.dll"

[Setup]
AppId={{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf32}\MistralOfficeAddin
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputBaseFilename=MistralOfficeAddin-Setup-x86
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x86 x64 arm64
ArchitecturesInstallIn64BitMode=

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; 32-bit binary payload
Source: "..\bin\x86\Release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Word Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\MistralAI.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Mistral AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\MistralAI.Addin"; ValueType: string; ValueName: "Description"; ValueData: "AI assistant using your own Mistral API key"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\MistralAI.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; Excel Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\MistralAI.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Mistral AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\MistralAI.Addin"; ValueType: string; ValueName: "Description"; ValueData: "AI assistant using your own Mistral API key"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\MistralAI.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; PowerPoint Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\MistralAI.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Mistral AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\MistralAI.Addin"; ValueType: string; ValueName: "Description"; ValueData: "AI assistant using your own Mistral API key"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\MistralAI.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; Outlook Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\MistralAI.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Mistral AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\MistralAI.Addin"; ValueType: string; ValueName: "Description"; ValueData: "AI assistant using your own Mistral API key"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Outlook\Addins\MistralAI.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; Task pane ActiveX (Office CreateCTP requires HKLM CLSID + Control categories)
Root: HKLM; Subkey: "Software\Classes\CLSID\{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Control"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\MiscStatus"; ValueType: string; ValueName: ""; ValueData: "131473"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Implemented Categories\{7DD95801-9882-11CF-9FA9-00AA006C42C4}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Implemented Categories\{7DD95802-9882-11CF-9FA9-00AA006C42C4}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}"; Flags: uninsdeletekey

[Run]
; Register COM Server via 32-bit RegAsm
Filename: "{dotnet4032}\regasm.exe"; Parameters: "/codebase ""{app}\MistralOfficeAddin.dll"""; StatusMsg: "Registering 32-bit COM components..."; Flags: runhidden

[UninstallRun]
; Unregister COM Server
Filename: "{dotnet4032}\regasm.exe"; Parameters: "/unregister ""{app}\MistralOfficeAddin.dll"""; StatusMsg: "Unregistering COM components..."; Flags: runhidden

[Code]
// Check .NET Framework 4.8 requirement
function IsDotNet48Installed(): Boolean;
var
  release: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', release) then
  begin
    // 528040 is release key for .NET Framework 4.8
    if release >= 528040 then
      Result := True;
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not IsDotNet48Installed() then
  begin
    if MsgBox('.NET Framework 4.8 or higher was not detected.' + #13#10 +
              'Would you like to open Microsoft''s download page to install it now?', mbConfirmation, MB_YESNO) = IDYES then
    begin
      ShellExec('open', 'https://dotnet.microsoft.com/download/dotnet-framework/net48', '', '', SW_SHOWNORMAL, ewNoWait, 0);
    end;
  end;
end;
