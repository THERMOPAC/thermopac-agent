; ThermopacAgent — Inno Setup 6 installer script
; Compile with: iscc setup.iss
; Requires: Inno Setup 6.x (https://jrsoftware.org/isinfo.php)
; Requires: dist\ThermopacAgent\ to exist (run build.bat first)

#define AppName      "ThermopacAgent"
#define AppVersion   "1.0.0"
#define AppPublisher "Thermopac"
#define AppURL       "https://thermopac-communication-thermopacllp.replit.app"
#define AppExeName   "ThermopacAgent.exe"
#define SourceDir    "..\dist\ThermopacAgent"

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
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";    Description: "Create a &desktop shortcut";       GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startupschedule"; Description: "Start agent automatically when &Windows logs in"; GroupDescription: "Auto-start:"; Flags: unchecked

[Files]
; All compiled files from PyInstaller dist folder
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Dirs]
; Create working directories with full user write permissions
Name: "{commonappdata}\ThermopacAgent\temp"; Permissions: everyone-full
Name: "{commonappdata}\ThermopacAgent\logs"; Permissions: everyone-full

[Icons]
Name: "{group}\{#AppName}";       Filename: "{app}\{#AppExeName}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent unchecked

[Registry]
; Task Scheduler auto-start on login (optional task)
Root: HKLM; Subkey: "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks"; Flags: dontcreatekey

[Code]
procedure CreateScheduledTask();
var
  ResultCode: Integer;
begin
  if IsTaskSelected('startupschedule') then
  begin
    Exec('schtasks.exe',
      '/Create /F /SC ONLOGON /RL HIGHEST /TN "ThermopacAgent" /TR "\"' +
      ExpandConstant('{app}') + '\ThermopacAgent.exe\"" /DELAY 0000:30',
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode = 0 then
      MsgBox('Scheduled task created: ThermopacAgent will start 30 seconds after login.',
             mbInformation, MB_OK)
    else
      MsgBox('Warning: Could not create scheduled task (exit code ' + IntToStr(ResultCode) + '). ' +
             'You can start the agent manually from the Start Menu.', mbError, MB_OK);
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    CreateScheduledTask();
    MsgBox('Installation complete!' + #13#10 + #13#10 +
           'Agent will auto-configure on first run:' + #13#10 +
           '  - node_id auto-filled from machine name' + #13#10 +
           '  - SolidWorks version auto-detected from registry' + #13#10 +
           '  - Testing mode: token auto-generated and self-registered' + #13#10 + #13#10 +
           'Production mode only:' + #13#10 +
           '  Set node_token in config.ini (admin-issued token required).' + #13#10 + #13#10 +
           'Launch ThermopacAgent from the Start Menu or Desktop shortcut.',
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
