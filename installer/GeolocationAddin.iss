; Inno Setup Script for Geolocation Addin (Revit 2024)
; Requires Inno Setup 6+ — https://jrsoftware.org/isinfo.php

#define MyAppName "Geolocation Addin"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Custom Geolocation Tools"

; Repo root is one level up from this script
#define RepoRoot ".."
#define BuildOutput RepoRoot + "\src\GeolocationAddin\bin\Release"
#define ConfigSamples RepoRoot + "\config"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={userappdata}\Autodesk\Revit\Addins\2024
DirExistsWarning=no
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=GeolocationAddin-Setup-{#MyAppVersion}
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
SetupIconFile=compiler:SetupClassicIcon.ico
UninstallDisplayName={#MyAppName} for Revit 2024

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; Add-in DLL, manifest, and dependencies -> Revit Addins folder
Source: "{#BuildOutput}\GeolocationAddin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\GeolocationAddin.addin"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#BuildOutput}\Newtonsoft.Json.dll"; DestDir: "{app}"; Flags: ignoreversion

; Sample config files -> ProgramData (don't overwrite existing)
Source: "{#ConfigSamples}\config.sample.json"; DestDir: "{commonappdata}\GeolocationAddin"; DestName: "config.json"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "{#ConfigSamples}\mapping.sample.csv"; DestDir: "{commonappdata}\GeolocationAddin"; DestName: "mapping.csv"; Flags: onlyifdoesntexist uninsneveruninstall

[Dirs]
Name: "{commonappdata}\GeolocationAddin"; Flags: uninsneveruninstall

[Messages]
SelectDirLabel3=The add-in will be installed into the following Revit 2024 add-ins folder.

[Code]
function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = wpSelectDir then
  begin
    { Warn if the selected directory doesn't look like a Revit addins folder }
    if Pos('Revit', WizardDirValue) = 0 then
    begin
      if MsgBox('The selected folder does not appear to be a Revit add-ins directory.' + #13#10 + #13#10 +
                'The default location is:' + #13#10 +
                ExpandConstant('{userappdata}') + '\Autodesk\Revit\Addins\2024' + #13#10 + #13#10 +
                'Continue anyway?', mbConfirmation, MB_YESNO) = IDNO then
        Result := False;
    end;
  end;
end;
