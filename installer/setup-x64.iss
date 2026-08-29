; Inno Setup 6 Script for AI Assistant Office Add-in (64-bit / x64)
; Target: 64-bit Microsoft Office (Word, Excel, PowerPoint)

#define MyAppName "MS Office AI Assistant (64-bit)"
#define MyAppVersion "0.6.0"
#define MyAppPublisher "D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302"
#define MyAppURL ""
#define MyAppExeName "MSOfficeAIAssistant.dll"

[Setup]
AppId={{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf64}\AIAssistant
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputBaseFilename=AIAssistant-Setup-x64
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesAllowed=x64 arm64
ArchitecturesInstallIn64BitMode=x64 arm64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; 64-bit binary payload
Source: "..\bin\x64\Release\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Word Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\MSOfficeAIAssistant.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "MS Office AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\MSOfficeAIAssistant.Addin"; ValueType: string; ValueName: "Description"; ValueData: "MS Office AI Assistant for Word, Excel, and PowerPoint"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\MSOfficeAIAssistant.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; Excel Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\MSOfficeAIAssistant.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "MS Office AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\MSOfficeAIAssistant.Addin"; ValueType: string; ValueName: "Description"; ValueData: "MS Office AI Assistant for Word, Excel, and PowerPoint"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Excel\Addins\MSOfficeAIAssistant.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; PowerPoint Add-in Registration
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\MSOfficeAIAssistant.Addin"; ValueType: string; ValueName: "FriendlyName"; ValueData: "MS Office AI Assistant"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\MSOfficeAIAssistant.Addin"; ValueType: string; ValueName: "Description"; ValueData: "MS Office AI Assistant for Word, Excel, and PowerPoint"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\PowerPoint\Addins\MSOfficeAIAssistant.Addin"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"; Flags: uninsdeletekey

; Task pane ActiveX (Office CreateCTP requires HKLM CLSID + Control categories)
; NOTE: every literal { in a quoted string is escaped as {{ — Inno Setup treats an
; unescaped { as the start of a {constant} reference (like {app} above) even inside
; quotes, and a bare CLSID/category GUID here is not a known constant, so it would
; fail to compile ("Unknown constant") without escaping every opening brace.
Root: HKLM; Subkey: "Software\Classes\CLSID\{{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Control"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\MiscStatus"; ValueType: string; ValueName: ""; ValueData: "131473"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Implemented Categories\{{7DD95801-9882-11CF-9FA9-00AA006C42C4}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Implemented Categories\{{7DD95802-9882-11CF-9FA9-00AA006C42C4}"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Classes\CLSID\{{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}\Implemented Categories\{{40FC6ED4-2438-11CF-A3DB-080036F12502}"; Flags: uninsdeletekey

[Run]
; Unregister any old COM registration first to avoid stale assembly version conflicts.
; This prevents manifest-mismatch COM activation failures (HRESULT 0x80131040) when
; upgrading from older build versions where assembly versions differed.
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/unregister ""{app}\MSOfficeAIAssistant.dll"""; StatusMsg: "Cleaning old COM registration..."; Flags: runhidden skipifsilent
; Register COM Server via 64-bit RegAsm
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/codebase ""{app}\MSOfficeAIAssistant.dll"""; StatusMsg: "Registering 64-bit COM components..."; Flags: runhidden

[UninstallRun]
; Unregister COM Server
Filename: "{dotnet4064}\regasm.exe"; Parameters: "/unregister ""{app}\MSOfficeAIAssistant.dll"""; StatusMsg: "Unregistering 64-bit COM components..."; Flags: runhidden

[Code]
// Check .NET Framework 4.8 requirement
function IsDotNet48Installed(): Boolean;
var
  release: Cardinal;
begin
  Result := False;
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', release) then
  begin
    if release >= 528040 then
      Result := True;
  end;
end;

function InitializeSetup(): Boolean;
var
  ResultCode: Integer;
begin
  Result := True;
  if not IsDotNet48Installed() then
  begin
    if MsgBox('.NET Framework 4.8 or higher was not detected.' + #13#10 +
              'Would you like to open Microsoft''s download page to install it now?', mbConfirmation, MB_YESNO) = IDYES then
    begin
      // ShellExec's final parameter is declared "var Integer" (an out param for the
      // process's exit code) — it must be a variable, not a literal; passing 0 fails
      // to compile with "Variable expected".
      ShellExec('open', 'https://dotnet.microsoft.com/download/dotnet-framework/net48', '', '', SW_SHOWNORMAL, ewNoWait, ResultCode);
    end;
  end;
end;
