; ============================================================================
;  Trial Quest PowerPoint Add-in  -  slim installer (Inno Setup 6)
; ----------------------------------------------------------------------------
;  Installs ONLY the small .ppam (+ registers it with PowerPoint), then pulls
;  the ~36 MB template assets from the private GitHub repo using a read-only
;  token entered during setup. The old monolithic exe bundled the assets; this
;  one downloads them, so the installer stays tiny and assets update over GitHub.
;
;  BUILD: open this in the Inno Setup Compiler (iscc.exe build\installer.iss).
;  Requires dist\TrialQuest.ppam (run build\build.ps1 first).
;
;  KEEP THESE IN SYNC with Updater.bas (GH_OWNER / GH_REPO / GH_BRANCH) and
;  with version.json (MyAppVersion == addinVersion).
; ============================================================================

#define MyAppName    "Trial Quest PowerPoint Add-in"
#define MyAppVersion "5.6.4"
#define MyPublisher  "Trial Quest"
#define MyAddinFile  "TrialQuest.ppam"
#define GhOwner      "RyanWW-Products"
#define GhRepo       "TQ-PPT-TOOLS"

; Channel: "stable" (default) or "beta". Build the BETA installer with:
;     ISCC.exe /DChannel=beta build\installer.iss
#ifndef Channel
  #define Channel "stable"
#endif
#if Channel == "stable"
  #define GhBranch "main"
  #define ChannelTag ""
#else
  #define GhBranch Channel
  #define ChannelTag " " + Channel
#endif

[Setup]
AppId={{B6E3A9F4-2C71-4D88-9A0E-5F2C7D1A4E33}
AppName={#MyAppName}{#ChannelTag}
AppVersion={#MyAppVersion}
AppPublisher={#MyPublisher}
; Per-user install into the PowerPoint add-ins folder (no admin needed).
PrivilegesRequired=lowest
DefaultDirName={userappdata}\Microsoft\AddIns\Trial Ex Addin
DisableDirPage=yes
DisableProgramGroupPage=yes
DisableReadyPage=no
Uninstallable=yes
OutputBaseFilename=TEI Addin Setup v{#MyAppVersion}{#ChannelTag}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
; SetupIconFile=..\Logos\TQ Icon.ico

[Files]
Source: "..\dist\{#MyAddinFile}"; DestDir: "{app}"; Flags: ignoreversion
Source: "download-assets.ps1";   DestDir: "{tmp}"; Flags: deleteafterinstall

[Registry]
; --- Update settings consumed by both the installer and the add-in ----------
Root: HKCU; Subkey: "Software\TrialQuest\Addin"; ValueType: string; ValueName: "GitHubToken"; ValueData: "{code:GetTokenInput}"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\TrialQuest\Addin"; ValueType: string; ValueName: "InstalledVersion"; ValueData: "{#MyAppVersion}"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\TrialQuest\Addin"; ValueType: string; ValueName: "Channel"; ValueData: "{#Channel}"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\TrialQuest";       Flags: uninsdeletekeyifempty
Root: HKCU; Subkey: "Software\TrialQuest\Addin"; Flags: uninsdeletekeyifempty

; --- Register the add-in for auto-load with the detected PowerPoint version --
; AutoLoad = 0xFFFFFFFF, Path = full path to the .ppam.
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\AddIns\TrialQuest"; ValueType: dword;  ValueName: "AutoLoad"; ValueData: "$ffffffff"; Flags: uninsdeletekey; Check: RegisterFor('16.0')
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\PowerPoint\AddIns\TrialQuest"; ValueType: string; ValueName: "Path";     ValueData: "{app}\{#MyAddinFile}"; Check: RegisterFor('16.0')
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\PowerPoint\AddIns\TrialQuest"; ValueType: dword;  ValueName: "AutoLoad"; ValueData: "$ffffffff"; Flags: uninsdeletekey; Check: RegisterFor('15.0')
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\PowerPoint\AddIns\TrialQuest"; ValueType: string; ValueName: "Path";     ValueData: "{app}\{#MyAddinFile}"; Check: RegisterFor('15.0')
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\PowerPoint\AddIns\TrialQuest"; ValueType: dword;  ValueName: "AutoLoad"; ValueData: "$ffffffff"; Flags: uninsdeletekey; Check: RegisterFor('14.0')
Root: HKCU; Subkey: "Software\Microsoft\Office\14.0\PowerPoint\AddIns\TrialQuest"; ValueType: string; ValueName: "Path";     ValueData: "{app}\{#MyAddinFile}"; Check: RegisterFor('14.0')

[UninstallDelete]
; Remove downloaded assets (not tracked by [Files]) and any update staging.
Type: filesandordirs; Name: "{app}"

[Code]
var
  TokenPage: TInputQueryWizardPage;

procedure InitializeWizard;
begin
  TokenPage := CreateInputQueryPage(wpSelectDir,
    'Update Access Token',
    'GitHub read-only token for updates',
    'Paste the read-only GitHub access token provided by your administrator. ' +
    'It is stored only on this machine and is used to download the template ' +
    'assets now and to fetch future updates from the "Check for Updates" button.');
  TokenPage.Add('Access token:', False);
end;

function GetTokenInput(Param: String): String;
begin
  Result := Trim(TokenPage.Values[0]);
end;

{ True if this Office version's PowerPoint is present. 16.0 also acts as the
  fallback when no Office version is detected at all. }
function PptPresent(ver: String): Boolean;
begin
  Result := RegKeyExists(HKCU, 'Software\Microsoft\Office\' + ver + '\PowerPoint') or
            RegKeyExists(HKLM, 'Software\Microsoft\Office\' + ver + '\PowerPoint');
end;

function RegisterFor(ver: String): Boolean;
begin
  if ver = '16.0' then
    Result := PptPresent('16.0') or (not PptPresent('15.0') and not PptPresent('14.0'))
  else
    Result := PptPresent(ver);
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  Params: String;
begin
  if CurStep = ssPostInstall then
  begin
    if GetTokenInput('') = '' then
    begin
      MsgBox('No token was entered, so template assets were not downloaded.' + #13#10 +
             'The add-in is installed; once a token is set you can fetch the ' +
             'assets from PowerPoint via Trial Quest > Check for Updates.',
             mbInformation, MB_OK);
      Exit;
    end;

    WizardForm.StatusLabel.Caption := 'Downloading template assets from GitHub...';
    Params := '-ExecutionPolicy Bypass -WindowStyle Hidden -File "' +
              ExpandConstant('{tmp}\download-assets.ps1') + '"' +
              ' -Owner "{#GhOwner}" -Repo "{#GhRepo}" -Branch "{#GhBranch}"' +
              ' -Dest "' + ExpandConstant('{userappdata}\Microsoft\AddIns') + '"';
    if not Exec('powershell.exe', Params, '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
      ResultCode := -1;

    if ResultCode <> 0 then
      MsgBox('Template assets could not be downloaded (code ' + IntToStr(ResultCode) + ').' + #13#10 +
             'The add-in is installed; you can retry later via ' +
             'Trial Quest > Check for Updates.', mbError, MB_OK);
  end;
end;
