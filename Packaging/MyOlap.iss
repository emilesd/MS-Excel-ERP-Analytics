; MyOlap - Excel OLAP Analytics Add-in
; Per-user installer: the add-in itself needs no administrator rights.
; Only the bundled .NET 8 Desktop Runtime prerequisite triggers elevation,
; and only when that runtime is not already present.
;
; Build:  ISCC.exe MyOlap.iss

#define AppName         "MyOlap"
#define AppVersion      "1.5.0"
#define AppPublisher    "GoLive Systems Ltd"
#define AppURL          "https://github.com/emilesd/MS-Excel-ERP-Analytics"
#define DotNetInstaller "windowsdesktop-runtime-8-win-x64.exe"

[Setup]
AppId={{8E1F2A64-3D7B-4C59-9A11-5C0F2D7E4B31}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppName}
VersionInfoCompany={#AppPublisher}
VersionInfoDescription={#AppName} Excel Add-in Setup

DefaultDirName={localappdata}\MyOlap\AddIn
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
DisableDirPage=yes
UsePreviousAppDir=yes

PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

OutputDir=output
OutputBaseFilename=MyOlap-Setup-{#AppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
UninstallDisplayName={#AppName} {#AppVersion} (Excel Add-in)
CloseApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
; --- Add-in payload: 64-bit Office ---
Source: "payload\MyOlap-AddIn64-packed.xll"; DestDir: "{app}"; Flags: ignoreversion; Check: IsOffice64
Source: "payload\e_sqlite3-x64.dll";         DestDir: "{app}"; DestName: "e_sqlite3.dll"; Flags: ignoreversion; Check: IsOffice64

; --- Add-in payload: 32-bit Office ---
Source: "payload\MyOlap-AddIn-packed.xll";   DestDir: "{app}"; Flags: ignoreversion; Check: not IsOffice64
Source: "payload\e_sqlite3-x86.dll";         DestDir: "{app}"; DestName: "e_sqlite3.dll"; Flags: ignoreversion; Check: not IsOffice64

; --- Runtime metadata required by the Excel-DNA .NET host ---
Source: "payload\MyOlap.deps.json";          DestDir: "{app}"; Flags: ignoreversion
Source: "payload\MyOlap-AddIn.deps.json";    DestDir: "{app}"; Flags: ignoreversion
Source: "payload\MyOlap-AddIn64.deps.json";  DestDir: "{app}"; Flags: ignoreversion
Source: "payload\MyOlap.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion

; --- Documentation ---
Source: "payload\UserGuide.md";      DestDir: "{app}"; Flags: ignoreversion
Source: "payload\Install-Notes.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme

; --- Prerequisite: extracted only when the runtime is missing ---
Source: "prereq\{#DotNetInstaller}"; DestDir: "{tmp}"; Flags: deleteafterinstall; Check: ShouldInstallDotNet

[Icons]
Name: "{group}\MyOlap User Guide"; Filename: "{app}\UserGuide.md"
Name: "{group}\MyOlap Install Notes"; Filename: "{app}\Install-Notes.txt"
Name: "{group}\Uninstall MyOlap"; Filename: "{uninstallexe}"

[Run]
Filename: "{tmp}\{#DotNetInstaller}"; Parameters: "/install /quiet /norestart"; \
  StatusMsg: "Installing .NET 8 Desktop Runtime (one-time; may prompt for administrator)..."; \
  Flags: shellexec waituntilterminated; Check: ShouldInstallDotNet

[Code]
const
  ADDIN_MARKER   = 'MyOlap-AddIn';
  MAX_OPEN_SLOTS = 30;
  DOTNET_URL     = 'https://dotnet.microsoft.com/download/dotnet/8.0';

var
  gOffice64: Boolean;
  gOffice64Known: Boolean;
  gExcelVerKey: String;
  gExcelVerKeyKnown: Boolean;

{ --------------------------------------------------------------------------
  Office bitness. Determines which packed XLL and which native SQLite build
  get installed, and which .NET runtime architecture Excel will need.
  -------------------------------------------------------------------------- }
function IsOffice64: Boolean;
var
  platform, installRoot: String;
begin
  if not gOffice64Known then
  begin
    if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\ClickToRun\Configuration',
                           'Platform', platform) then
      gOffice64 := (CompareText(Trim(platform), 'x64') = 0)
    else if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot',
                                'Path', installRoot) then
      gOffice64 := (Pos('program files (x86)', Lowercase(installRoot)) = 0)
    else
      gOffice64 := True;

    gOffice64Known := True;
  end;
  Result := gOffice64;
end;

function XllFileName: String;
begin
  if IsOffice64 then
    Result := 'MyOlap-AddIn64-packed.xll'
  else
    Result := 'MyOlap-AddIn-packed.xll';
end;

{ --------------------------------------------------------------------------
  .NET 8 Desktop Runtime detection.
  The dotnet registry keys are not written by every installer flavour, so the
  shared-framework folder is the reliable signal.
  -------------------------------------------------------------------------- }
function HasDotNet8DesktopIn(const DotNetRoot: String): Boolean;
var
  rec: TFindRec;
  base: String;
begin
  Result := False;
  base := DotNetRoot + '\shared\Microsoft.WindowsDesktop.App';
  if not DirExists(base) then
    exit;

  if FindFirst(base + '\*', rec) then
  try
    repeat
      if ((rec.Attributes and FILE_ATTRIBUTE_DIRECTORY) <> 0)
         and (Copy(rec.Name, 1, 2) = '8.') then
        Result := True;
    until Result or (not FindNext(rec));
  finally
    FindClose(rec);
  end;
end;

function DotNet8Present: Boolean;
var
  root: String;
begin
  { Must be the machine-wide Program Files, never the per-user "auto" variant:
    the .NET shared runtime is always installed machine-wide. }
  if IsOffice64 then
  begin
    root := GetEnv('ProgramW6432');
    if root = '' then
      root := ExpandConstant('{commonpf64}');
  end
  else
  begin
    root := GetEnv('ProgramFiles(x86)');
    if root = '' then
      root := ExpandConstant('{commonpf32}');
  end;

  Result := HasDotNet8DesktopIn(root + '\dotnet');
end;

{ The bundled prerequisite is the x64 runtime, so it only helps 64-bit Office. }
function ShouldInstallDotNet: Boolean;
begin
  Result := IsOffice64 and (not DotNet8Present);
end;

{ --------------------------------------------------------------------------
  Excel registry version key: 16.0 covers Office 2016/2019/2021/365.
  -------------------------------------------------------------------------- }
function ExcelVersionKey: String;
var
  candidates: array[0..2] of String;
  i: Integer;
  dummy: String;
begin
  if not gExcelVerKeyKnown then
  begin
    candidates[0] := '16.0';
    candidates[1] := '15.0';
    candidates[2] := '14.0';

    gExcelVerKey := '';
    for i := 0 to 2 do
      if gExcelVerKey = '' then
        if RegQueryStringValue(HKLM64, 'SOFTWARE\Microsoft\Office\' + candidates[i]
                               + '\Excel\InstallRoot', 'Path', dummy)
           or RegKeyExists(HKCU, 'Software\Microsoft\Office\' + candidates[i]
                           + '\Excel\Options') then
          gExcelVerKey := candidates[i];

    if gExcelVerKey = '' then
      gExcelVerKey := '16.0';

    gExcelVerKeyKnown := True;
  end;
  Result := gExcelVerKey;
end;

function ExcelOptionsKey: String;
begin
  Result := 'Software\Microsoft\Office\' + ExcelVersionKey + '\Excel\Options';
end;

function OpenSlotName(Index: Integer): String;
begin
  if Index = 0 then
    Result := 'OPEN'
  else
    Result := 'OPEN' + IntToStr(Index);
end;

{ --------------------------------------------------------------------------
  Excel auto-loads the add-ins listed in OPEN, OPEN1, OPEN2... and stops at the
  first missing slot, so the list must stay contiguous.
  -------------------------------------------------------------------------- }
procedure RegisterAddIn;
var
  key, data, target: String;
  i, slot: Integer;
begin
  key := ExcelOptionsKey;
  target := '/R "' + ExpandConstant('{app}') + '\' + XllFileName + '"';

  slot := -1;
  for i := 0 to MAX_OPEN_SLOTS - 1 do
  begin
    if RegQueryStringValue(HKCU, key, OpenSlotName(i), data) then
    begin
      if Pos(Lowercase(ADDIN_MARKER), Lowercase(data)) > 0 then
      begin
        slot := i;
        break;
      end;
    end
    else
    begin
      slot := i;
      break;
    end;
  end;

  if slot < 0 then
    slot := 0;

  RegWriteStringValue(HKCU, key, OpenSlotName(slot), target);
end;

procedure UnregisterAddIn;
var
  key, data: String;
  kept: array[0..MAX_OPEN_SLOTS] of String;
  i, count: Integer;
begin
  key := ExcelOptionsKey;
  count := 0;

  for i := 0 to MAX_OPEN_SLOTS - 1 do
  begin
    if not RegQueryStringValue(HKCU, key, OpenSlotName(i), data) then
      break;
    if Pos(Lowercase(ADDIN_MARKER), Lowercase(data)) = 0 then
    begin
      kept[count] := data;
      count := count + 1;
    end;
    RegDeleteValue(HKCU, key, OpenSlotName(i));
  end;

  for i := 0 to count - 1 do
    RegWriteStringValue(HKCU, key, OpenSlotName(i), kept[i]);
end;

{ Excel parks add-ins that once crashed in a "disabled items" list; clearing it
  lets a fresh install load instead of silently doing nothing. }
procedure ClearExcelDisabledItems;
begin
  RegDeleteKeyIncludingSubkeys(HKCU,
    'Software\Microsoft\Office\' + ExcelVersionKey + '\Excel\Resiliency\DisabledItems');
end;

function IsExcelRunning: Boolean;
var
  code: Integer;
begin
  Result := False;
  if Exec(ExpandConstant('{cmd}'),
          '/C tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE"',
          '', SW_HIDE, ewWaitUntilTerminated, code) then
    Result := (code = 0);
end;

function EnsureExcelClosed(const Action: String): Boolean;
var
  code: Integer;
begin
  Result := True;
  if not IsExcelRunning then
    exit;

  if MsgBox('Excel is running and must be closed before MyOlap can be ' + Action + '.'
            + #13#10#13#10 + 'Save your work, then click Yes to close Excel now.',
            mbConfirmation, MB_YESNO) = IDYES then
  begin
    Exec(ExpandConstant('{cmd}'), '/C taskkill /F /IM EXCEL.EXE', '', SW_HIDE,
         ewWaitUntilTerminated, code);
    Sleep(1500);
    Result := not IsExcelRunning;
  end
  else
    Result := False;

  if not Result then
    MsgBox('Excel is still running. Please close Excel and try again.', mbError, MB_OK);
end;

function InitializeSetup: Boolean;
begin
  Result := EnsureExcelClosed('installed');
end;

function InitializeUninstall: Boolean;
begin
  Result := EnsureExcelClosed('removed');
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  errCode: Integer;
begin
  if CurStep <> ssPostInstall then
    exit;

  ClearExcelDisabledItems;
  RegisterAddIn;

  if DotNet8Present or WizardSilent then
    exit;

  if IsOffice64 then
    MsgBox('MyOlap needs the .NET 8 Desktop Runtime (x64), which could not be installed '
           + 'automatically - administrator rights are required.' + #13#10#13#10
           + 'Install "Desktop Runtime 8.0 - x64" from the page that opens next, '
           + 'then start Excel.', mbError, MB_OK)
  else
    MsgBox('This machine has 32-bit Excel, so MyOlap needs the .NET 8 Desktop Runtime (x86), '
           + 'which is not bundled with this installer.' + #13#10#13#10
           + 'Install "Desktop Runtime 8.0 - x86" from the page that opens next, '
           + 'then start Excel.', mbError, MB_OK);

  ShellExec('open', DOTNET_URL, '', '', SW_SHOWNORMAL, ewNoWait, errCode);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    UnregisterAddIn;
end;
