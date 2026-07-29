; AutoOffice Word Add-in Installer
; Inno Setup Script
; https://jrsoftware.org/isinfo.php

#define MyAppName "AutoOffice for Word, Excel & PowerPoint"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "AutoOffice"
#define MyAppURL "https://sivan22.github.io/autoOffice/"
#define ShareName "AutoOfficeAddin"

[Setup]
AppId={{B2C3D4E5-F6A7-8901-BCDE-F12345678902}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
DefaultDirName={commonappdata}\AutoOfficeAddin
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=AutoOffice-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"

[Files]
Source: "..\manifest.production.xml"; DestDir: "{app}"; DestName: "manifest.xml"; Flags: ignoreversion
Source: "..\apps\server\dist\autoOffice-server.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "launcher.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\apps\web\dist\*"; DestDir: "{app}\web"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
; Only register our own trusted catalog when none already exists. Office's
; per-user TrustedCatalogs parser breaks with 2+ GUID-named subkeys: it shows
; "we had a problem reading your settings" and wipes ALL entries on Word
; startup. When a catalog already exists we drop our manifest into its folder
; instead, leaving the host catalog as the single registered entry.
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{B2C3D4E5-F6A7-8901-BCDE-F12345678903}"; ValueType: string; ValueName: "Id"; ValueData: "{{B2C3D4E5-F6A7-8901-BCDE-F12345678903}"; Flags: uninsdeletekey; Check: ShouldCreateOwnCatalog
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{B2C3D4E5-F6A7-8901-BCDE-F12345678903}"; ValueType: string; ValueName: "Url"; ValueData: "{code:GetNetworkPath}"; Check: ShouldCreateOwnCatalog
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{{B2C3D4E5-F6A7-8901-BCDE-F12345678903}"; ValueType: dword; ValueName: "Flags"; ValueData: "1"; Check: ShouldCreateOwnCatalog

[Code]
var
  NetworkPath: string;
  OwnSharePath: string;
  HostCatalogUrl: string;
  UseHostCatalog: Boolean;
  CreatedOwnShare: Boolean;

function GetNetworkPath(Param: string): string;
begin
  Result := NetworkPath;
end;

function GetComputerNetName: string;
begin
  Result := GetEnv('COMPUTERNAME');
  if Result = '' then
    Result := 'localhost';
end;

function CreateNetworkShare(ShareName, SharePath: string): Boolean;
var
  ResultCode: Integer;
  Command: string;
  CurrentUser: string;
begin
  CurrentUser := GetEnv('USERNAME');
  if CurrentUser = '' then
  begin
    Result := False;
    Exit;
  end;
  // The Office catalog only needs to read its manifest. Never expose a
  // world-writable share, and never delete a pre-existing share with the same
  // name: a name collision must fail closed.
  Command :=
    'share ' + ShareName + '="' + SharePath + '" /GRANT:"' +
    CurrentUser + '",READ';
  Result := Exec('net', Command, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) and (ResultCode = 0);
end;

// Return the Url of the first existing UNC-path trusted catalog, or '' if none.
function FindHostCatalogUrl: string;
var
  KeyPath: string;
  Subkeys: TArrayOfString;
  i: Integer;
  Url: string;
begin
  Result := '';
  KeyPath := 'Software\Microsoft\Office\16.0\WEF\TrustedCatalogs';
  if not RegGetSubKeyNames(HKCU, KeyPath, Subkeys) then
    Exit;
  for i := 0 to GetArrayLength(Subkeys) - 1 do
  begin
    if RegQueryStringValue(HKCU, KeyPath + '\' + Subkeys[i], 'Url', Url) then
    begin
      if (Length(Url) > 2) and (Url[1] = '\') and (Url[2] = '\') then
      begin
        Result := Url;
        Exit;
      end;
    end;
  end;
end;

procedure InitializeWizard;
begin
  OwnSharePath := ExpandConstant('{commonappdata}\AutoOfficeAddin\catalog');
  NetworkPath := '\\' + GetComputerNetName + '\{#ShareName}';
  HostCatalogUrl := FindHostCatalogUrl;
  UseHostCatalog := HostCatalogUrl <> '';
  CreatedOwnShare := False;
end;

function ShouldCreateOwnCatalog: Boolean;
begin
  Result := not UseHostCatalog;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  TargetFile: string;
begin
  // Create only AutoOffice-owned resources. We intentionally do not kill
  // Office applications or clear the shared WEF cache: either action can lose
  // unsaved work or remove state belonging to other add-ins.
  if CurStep = ssInstall then
  begin
    if not UseHostCatalog then
    begin
      if not DirExists(OwnSharePath) then
        CreateDir(OwnSharePath);
      CreatedOwnShare := CreateNetworkShare('{#ShareName}', OwnSharePath);
      if not CreatedOwnShare then
      begin
        MsgBox(
          'AutoOffice could not create its read-only Office catalog share. ' +
          'A share named "{#ShareName}" may already exist. Installation will stop without changing that share.',
          mbError,
          MB_OK
        );
        Abort;
      end;
      RegWriteDWordValue(HKCU, 'Software\AutoOffice\Installer', 'OwnsCatalogShare', 1);
    end;
  end;
  if CurStep = ssPostInstall then
  begin
    if UseHostCatalog then
      // Distinct filename so we never overwrite the host catalog's own manifest.
      TargetFile := HostCatalogUrl + '\autooffice.xml'
    else
      TargetFile := '{#OwnSharePath}\manifest.xml';
    if not CopyFile(ExpandConstant('{app}\manifest.xml'), TargetFile, False) then
      MsgBox('Warning: Could not copy manifest to share folder: ' + TargetFile, mbInformation, MB_OK);
    // Record where we placed the manifest so the uninstaller can remove it.
    RegWriteStringValue(HKCU, 'Software\AutoOffice\Installer', 'ManifestPath', TargetFile);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
  ManifestPath: string;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    OwnSharePath := ExpandConstant('{commonappdata}\AutoOfficeAddin\catalog');
    // Remove only the manifest we placed; never touch a catalog/share we
    // didn't create.
    if RegQueryStringValue(HKCU, 'Software\AutoOffice\Installer', 'ManifestPath', ManifestPath) then
    begin
      if FileExists(ManifestPath) then
        DeleteFile(ManifestPath);
    end;
    // If the standalone share/folder exists, it was created by us.
    if RegValueExists(HKCU, 'Software\AutoOffice\Installer', 'OwnsCatalogShare') then
    begin
      Exec('net', 'share {#ShareName} /delete /y', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
      if DirExists(OwnSharePath) then
        DelTree(OwnSharePath, True, True, True);
      RegDeleteValue(HKCU, 'Software\AutoOffice\Installer', 'OwnsCatalogShare');
    end;
    if MsgBox('Also remove your AutoOffice data folder (chat history, settings, provider keys)?', mbConfirmation, MB_YESNO) = IDYES then
    begin
      DelTree(ExpandConstant('{localappdata}\AutoOffice'), True, True, True);
    end;
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\AutoOffice');
  end;
end;

[Run]
Filename: "{app}\autoOffice-server.exe"; Parameters: "--first-run-init"; Flags: runhidden waituntilterminated; StatusMsg: "Initializing AutoOffice (cert + token) ..."
Filename: "schtasks.exe"; Parameters: "/Create /F /SC ONLOGON /TN ""AutoOffice\Service"" /TR ""wscript.exe //B //NoLogo {app}\launcher.vbs"" /RL LIMITED"; Flags: runhidden waituntilterminated; StatusMsg: "Registering AutoOffice Service ..."
Filename: "schtasks.exe"; Parameters: "/Run /TN ""AutoOffice\Service"""; Flags: runhidden waituntilterminated; StatusMsg: "Starting AutoOffice Service ..."

[Messages]
english.FinishedLabel=Installation completed successfully.%n%nTo use the add-in in Word, Excel, or PowerPoint:%n%n1. Open Microsoft Word, Excel, or PowerPoint%n2. Go to Home > Add-ins%n3. Select "Shared Folder" at the bottom%n4. Choose "AutoOffice" and click Add
hebrew.FinishedLabel=ההתקנה הסתיימה בהצלחה.%n%nכדי להשתמש בתוסף ב-Word, Excel או PowerPoint:%n%n1. פתח את Microsoft Word, Excel או PowerPoint%n2. עבור לעמוד הבית > תוספות%n3. לחץ על "תיקייה משותפת" בחלק התחתון%n4. בחר "AutoOffice" והקלק הוסף

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[UninstallRun]
Filename: "schtasks.exe"; Parameters: "/End /TN ""AutoOffice\Service"""; Flags: runhidden; RunOnceId: "stoptask"
Filename: "schtasks.exe"; Parameters: "/Delete /F /TN ""AutoOffice\Service"""; Flags: runhidden; RunOnceId: "deltask"
Filename: "{app}\autoOffice-server.exe"; Parameters: "--cert-uninstall"; Flags: runhidden; RunOnceId: "certrm"
