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

#define AppName      "ThermopacAgent Dev"
#define AppVersion   "1.0.60"
#define AppPublisher "Thermopac"
#define AppURL       "https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev"
#define AppExeName   "run.bat"
; SourcePath is an InnoSetup built-in: the directory containing this .iss file
; (always local-agent\installer\). Using it avoids CWD-relative path bugs
; when ISCC is invoked from a different working directory (e.g. the repo root
; or local-agent\ in GitHub Actions).
#define AgentRoot    SourcePath + "\.."
#define SourceDir    AgentRoot + "\dist\ThermopacAgent"
#define PythonDir    AgentRoot + "\dist\python"

[Setup]
AppId={{B06B4460-8C38-4705-97DC-71E9C1E5D936}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
DefaultDirName={autopf}\ThermopacAgentDev
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
; Agent EXE + supporting files (PyInstaller output)
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Bundled Python runtime — optional.  If dist\python\ was not created at build time
; (e.g. a CI run without the embeddable download step) ISCC skips this entry safely.
; When present the installer is fully self-contained; when absent setup.ps1 downloads
; Python from python.org at install time.
Source: "{#PythonDir}\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist
; Launch scripts (not part of PyInstaller output — must be listed explicitly)
Source: "{#SourcePath}run.bat";           DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}run-service.bat";   DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}makepy-repair.bat"; DestDir: "{app}"; Flags: ignoreversion
; PowerShell bootstrap — always bundled so Python can be fetched if not pre-bundled
Source: "{#SourcePath}setup.ps1"; DestDir: "{app}"; Flags: ignoreversion
; Default config — users edit this after installation
Source: "{#AgentRoot}\config.ini"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

[Dirs]
Name: "{commonappdata}\ThermopacAgentDev\temp";   Permissions: everyone-full
Name: "{commonappdata}\ThermopacAgentDev\logs";   Permissions: everyone-full
Name: "{commonappdata}\ThermopacAgentDev\config"; Permissions: everyone-full

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
      '/Create /F /SC ONLOGON /RL HIGHEST /TN "ThermopacAgentDev" /TR "\"' +
      ExpandConstant('{app}') + '\run-service.bat\"" /DELAY 0000:30',
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode = 0 then
      MsgBox('Scheduled task created: ThermopacAgentDev starts 30s after login.',
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
  PyExe, Ps1: String;
  ResultCode: Integer;
begin
  if CurStep = ssPostInstall then
  begin
    PyExe := ExpandConstant('{app}\python\python.exe');
    Ps1   := ExpandConstant('{app}\setup.ps1');

    // If Python was NOT bundled by the installer, run setup.ps1 to fetch it
    if not FileExists(PyExe) then
    begin
      if FileExists(Ps1) then
      begin
        MsgBox(
          'Python was not bundled in this installer.' + #13#10 +
          'setup.ps1 will now download Python 3.11 from python.org.' + #13#10 + #13#10 +
          'A PowerShell window will open — leave it running until complete.',
          mbInformation, MB_OK);
        Exec('powershell.exe',
          '-ExecutionPolicy Bypass -File "' + Ps1 + '"',
          ExpandConstant('{app}'), SW_SHOW, ewWaitUntilTerminated, ResultCode);
      end else
        MsgBox(
          'Python not found and setup.ps1 is missing.' + #13#10 +
          'Run "Repair COM Cache" from the Start Menu after placing setup.ps1 in ' +
          ExpandConstant('{app}') + '.', mbError, MB_OK);
    end else
    begin
      // Python IS bundled — run makepy for the detected SolidWorks version
      ProgId := DetectSwProgId();
      if ProgId <> '' then
        RunMakepy(ProgId);
    end;

    CreateScheduledTask();

    MsgBox(
      'Installation complete!' + #13#10 + #13#10 +
      'NEXT STEPS:' + #13#10 +
      '  1. This Dev build installs side-by-side in ThermopacAgentDev' + #13#10 +
      '     api_url is set to the Development backend' + #13#10 +
      '     Testing mode: node_token is auto-generated if missing' + #13#10 +
      '     Production mode: paste your admin-issued token' + #13#10 + #13#10 +
      '  2. Run ThermopacAgent from the Start Menu' + #13#10 + #13#10 +
      '  3. If SolidWorks COM errors appear, run:' + #13#10 +
      '     Start Menu > ThermopacAgent Dev > Repair COM Cache',
      mbInformation, MB_OK);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
    Exec('schtasks.exe', '/Delete /F /TN "ThermopacAgentDev"',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
