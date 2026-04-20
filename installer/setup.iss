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
#define AppVersion   "1.0.20"
#define AppPublisher "Thermopac"
#define AppURL       "https://thermopac-communication-thermopacllp.replit.app"
#define AppExeName   "run.bat"
; SourcePath is an InnoSetup built-in: the directory containing this .iss file
; (always local-agent\installer\). Using it avoids CWD-relative path bugs
; when ISCC is invoked from a different working directory (e.g. the repo root
; or local-agent\ in GitHub Actions).
#define AgentRoot    SourcePath + "\.."
#define SourceDir    AgentRoot + "\dist\ThermopacAgent"
#define PythonDir    AgentRoot + "\dist\python"

[Setup]
AppId={{7B4A1C2D-E3F5-4A6B-8C9D-0E1F2A3B4C5D}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir={#AgentRoot}\installer_output
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

// ── SolidWorks detection ──────────────────────────────────────────────────
function DetectSwProgId(): String;
var
  Years: array of String;
  i: Integer;
  Dummy: String;
begin
  Result := '';
  Years := ['32','31','30','29','28','27'];
  for i := 0 to GetArrayLength(Years) - 1 do
  begin
    if RegQueryStringValue(HKCR, 'SldWorks.Application.' + Years[i], '', Dummy) then
    begin
      Result := 'SldWorks.Application.' + Years[i];
      Exit;
    end;
  end;
end;

function CheckSolidWorks(): Boolean;
begin
  Result := DetectSwProgId() <> '';
end;

// ── makepy: generate win32com early-binding cache ─────────────────────────
procedure RunMakepy(ProgId: String);
var
  PyExe, Args: String;
  ResultCode: Integer;
begin
  PyExe := ExpandConstant('{app}\python\python.exe');
  if not FileExists(PyExe) then
  begin
    MsgBox('Python not found at ' + PyExe + ' — skipping COM cache setup.' + #13#10 +
           'Run "Repair COM Cache" from the Start Menu after installation.',
           mbError, MB_OK);
    Exit;
  end;

  // Method A: standard makepy via ProgID
  Args := '-m win32com.client.makepy "' + ProgId + '"';
  Exec(PyExe, Args, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if ResultCode = 0 then
    Exit;

  // Method B: inline pythoncom registry walk (same logic as _prepare_sw_makepy_cache)
  Args := '-c "import winreg,pythoncom,win32com.client.gencache as g,sys; ' +
          'p=sys.argv[1]; ' +
          'k=winreg.OpenKey(winreg.HKEY_CLASSES_ROOT,p+chr(92)+chr(67)+chr(76)+chr(83)+chr(73)+chr(68)); ' +
          'c=winreg.QueryValue(k,chr(0)); winreg.CloseKey(k); ' +
          'print(chr(79)+chr(75))" "' + ProgId + '"';
  Exec(PyExe, Args, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  // Non-fatal — agent has its own 3-method fallback at runtime
end;

// ── Scheduled task ────────────────────────────────────────────────────────
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
             'Start from the Start Menu manually.', mbError, MB_OK);
  end;
end;

// ── Validation ────────────────────────────────────────────────────────────
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

// ── Post-install ──────────────────────────────────────────────────────────
procedure CurStepChanged(CurStep: TSetupStep);
var
  ProgId: String;
begin
  if CurStep = ssPostInstall then
  begin
    // Generate SolidWorks COM type library cache (makepy) in the bundled Python
    ProgId := DetectSwProgId();
    if ProgId <> '' then
      RunMakepy(ProgId);

    CreateScheduledTask();

    MsgBox(
      'Installation complete!' + #13#10 + #13#10 +
      'NEXT STEPS:' + #13#10 +
      '  1. Edit config.ini in the installation folder' + #13#10 +
      '     Set api_url to your Thermopac ERP URL' + #13#10 +
      '     Testing mode: node_token is auto-generated' + #13#10 +
      '     Production mode: paste your admin-issued token' + #13#10 + #13#10 +
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
