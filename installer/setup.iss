; ThermopacAgent — Inno Setup 6 installer script
; Compile with: iscc setup.iss   (from the installer\ directory)
;
; Build pipeline:
;   1. build-installer.bat        — downloads Python embeddable, pip install, makepy
;   2. iscc installer\setup.iss   — packages everything into a single .exe installer
;
; Requires:
;   - Inno Setup 6.x  https://jrsoftware.org/isinfo.php
;   - dist\ThermopacAgent\  folder built by build-installer.bat
;   - dist\python\          Python 3.11 embeddable (downloaded by build-installer.bat)

#define AppName      "ThermopacAgent"
#define AppVersion   "1.0.19"
#define AppPublisher "Thermopac"
#define AppURL       "https://thermopac-communication-thermopacllp.replit.app"
#define AppExeName   "run.bat"
#define SourceDir    "..\dist\ThermopacAgent"
#define PythonDir    "..\dist\python"

[Setup]
AppId={{7B4A1C2D-E3F5-4A6B-8C9D-0E1F2A3B4C5D}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=..\installer_output
OutputBaseFilename=ThermopacAgent-Setup-v{#AppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\run.bat
UninstallDisplayName={#AppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";      Description: "Create a &desktop shortcut";               GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startupschedule";  Description: "Start agent automatically at &Windows login"; GroupDescription: "Auto-start:";       Flags: unchecked

[Files]
; Agent source files
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Bundled Python runtime (includes makepy cache generated at build time)
Source: "{#PythonDir}\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
Name: "{commonappdata}\ThermopacAgent\temp";   Permissions: everyone-full
Name: "{commonappdata}\ThermopacAgent\logs";   Permissions: everyone-full
Name: "{commonappdata}\ThermopacAgent\config"; Permissions: everyone-full

[Icons]
Name: "{group}\{#AppName}";             Filename: "{app}\run.bat";       WorkingDir: "{app}"
Name: "{group}\Edit Config";            Filename: "notepad.exe";         Parameters: """{app}\config.ini"""
Name: "{group}\Repair COM Cache";       Filename: "{app}\makepy-repair.bat"; WorkingDir: "{app}"
Name: "{group}\Uninstall {#AppName}";   Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";       Filename: "{app}\run.bat";       WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\run.bat"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent unchecked shellexec

[Code]
procedure CreateScheduledTask();
var
  ResultCode: Integer;
begin
  if IsTaskSelected('startupschedule') then
  begin
    Exec('schtasks.exe',
      '/Create /F /SC ONLOGON /RL HIGHEST /TN "ThermopacAgent" /TR "\"' +
      ExpandConstant('{app}') + '\run-service.bat\"" /DELAY 0000:30',
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode = 0 then
      MsgBox('Scheduled task created: ThermopacAgent starts 30s after login.',
             mbInformation, MB_OK)
    else
      MsgBox('Warning: Could not create scheduled task (code ' + IntToStr(ResultCode) + '). ' +
             'Run from Start Menu manually.', mbError, MB_OK);
  end;
end;

function CheckSolidWorks(): Boolean;
var
  Years: array of String;
  i: Integer;
  Dummy: String;
begin
  Result := False;
  Years := ['32','31','30','29','28','27'];
  for i := 0 to GetArrayLength(Years) - 1 do
  begin
    if RegQueryStringValue(HKCR, 'SldWorks.Application.' + Years[i], '', Dummy) then
    begin
      Result := True;
      Exit;
    end;
  end;
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if not CheckSolidWorks() then
  begin
    MsgBox(
      'SolidWorks (2019-2024) was not detected on this machine.' + #13#10 + #13#10 +
      'ThermopacAgent requires SolidWorks to be installed.' + #13#10 +
      'Please install SolidWorks first, then re-run this installer.' + #13#10 + #13#10 +
      'Click OK to cancel installation.',
      mbCriticalError, MB_OK);
    Result := False;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    CreateScheduledTask();
    MsgBox(
      'Installation complete!' + #13#10 + #13#10 +
      'NEXT STEPS:' + #13#10 +
      '  1. Edit config.ini in the installation folder' + #13#10 +
      '     Set api_url to your Thermopac ERP URL' + #13#10 +
      '     In testing mode: node_token is auto-generated' + #13#10 +
      '     In production mode: paste your admin-issued token' + #13#10 + #13#10 +
      '  2. Run ThermopacAgent from the Start Menu' + #13#10 + #13#10 +
      '  3. If SolidWorks COM errors appear, run:' + #13#10 +
      '     Start Menu > ThermopacAgent > Repair COM Cache',
      mbInformation, MB_OK);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
    Exec('schtasks.exe', '/Delete /F /TN "ThermopacAgent"',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
