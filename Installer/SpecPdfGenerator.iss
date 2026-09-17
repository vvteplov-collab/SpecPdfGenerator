#define MyAppName "SpecPdfGenerator"
#define MyAppVersion "2.0"
#define MyAppPublisher "ОВиК"
#define MyAppExeName "SpecPdfGenerator.exe"
#define PublishDir "..\bin\Release\net8.0-windows\win-x64\publish"
#define AddinFile "..\ExcelAddin\SpecPdfGenerator.xlam"

[Setup]
AppId={{F82769B0-E024-443D-AE8A-A2AD2C62B567}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
UninstallDisplayName={#MyAppName}
OutputDir=Output
OutputBaseFilename=SpecPdfGenerator-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; The executable is installed per-machine, while the Excel add-in is intentionally per-user.
UsedUserAreasWarning=no
WizardStyle=modern

[Files]
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Excel automatically loads XLAM files from this per-user startup folder.
Source: "{#AddinFile}"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; DestName: "SpecPdfGenerator.xlam"; Flags: ignoreversion

[UninstallDelete]
; Remove only the XLAM belonging to this product.
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\SpecPdfGenerator.xlam"

[Code]
const
  ExcelOptionsKey = 'Software\Microsoft\Office\16.0\Excel\Options';
  LegacyExcelOptionsKey = 'Software\Microsoft\Office\Excel\Options';
  AddinSettingsKey = 'Software\SpecPdfGenerator';
  VbaSettingsKey = 'Software\VB and VBA Program Settings\SpecPdfGenerator\Settings';

procedure UnregisterAddinFromKey(const OptionsKey, AddinValue: String);
var
  Index: Integer;
  Name, Existing: String;
begin
  for Index := 0 to 99 do
  begin
    if Index = 0 then Name := 'OPEN' else Name := 'OPEN' + IntToStr(Index);
    if RegQueryStringValue(HKCU, OptionsKey, Name, Existing) and
       (CompareText(Existing, AddinValue) = 0) then
      RegDeleteValue(HKCU, OptionsKey, Name);
  end;
end;

procedure RemoveLegacyOpenRegistrations;
var
  AddinValue: String;
begin
  { Versions prior to 1.6 registered the XLAM through Excel\Options\OPEN*. }
  AddinValue := '/R "' + ExpandConstant('{app}\SpecPdfGenerator.xlam') + '"';
  UnregisterAddinFromKey(ExcelOptionsKey, AddinValue);
  UnregisterAddinFromKey(LegacyExcelOptionsKey, AddinValue);
  RegDeleteValue(HKCU, AddinSettingsKey, 'ExcelOpenValueName');
end;

procedure SetDefaultGeneratorPath;
var
  Existing: String;
begin
  if not RegQueryStringValue(HKCU, VbaSettingsKey, 'ExePath', Existing) or not FileExists(Existing) then
    RegWriteStringValue(HKCU, VbaSettingsKey, 'ExePath', ExpandConstant('{app}\SpecPdfGenerator.exe'));
end;

function InitializeSetup(): Boolean;
begin
  MsgBox('Перед установкой закройте Excel. После установки запустите Excel заново, чтобы загрузилась вкладка Teplov.', mbInformation, MB_OK);
  Result := True;
end;

procedure RemoveLegacyTrustedLocation;
var
  TrustedKey: String;
begin
  { XLSTART no longer needs the Program Files trusted location. }
  if RegQueryStringValue(HKCU, AddinSettingsKey, 'TrustedLocationKey', TrustedKey) then
    RegDeleteKeyIncludingSubkeys(HKCU, TrustedKey);
  RegDeleteValue(HKCU, AddinSettingsKey, 'TrustedLocationKey');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    RemoveLegacyOpenRegistrations;
    RemoveLegacyTrustedLocation;
    SetDefaultGeneratorPath;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
  begin
    RemoveLegacyOpenRegistrations;
    RemoveLegacyTrustedLocation;
  end;
end;
