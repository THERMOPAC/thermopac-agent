; ============================================================
; Thermopac Drawing Structuring Agent — Inno Setup 6 script
; Version: read from STRUCTURER_VERSION env var at compile time
;
; Compile with:
;   iscc installer\setup.iss                (from structurer_pkg\ directory)
;   OR
;   installer\build-installer.bat           (full pipeline — builds dist first)
;   OR (CI)
;   set STRUCTURER_VERSION=1.0.24 && iscc installer\setup.iss
;
; Prerequisites:
;   - Inno Setup 6.x  https://jrsoftware.org/isinfo.php
;   - dist\ThermopacStructuringAgent\  built by build-installer.bat
;   - dist\python\                     bundled Python (built by build-installer.bat)
; ============================================================

#define MyAppVersion        GetEnv("STRUCTURER_VERSION")
#if MyAppVersion == ""
  #define MyAppVersion      "1.0.24"
#endif

#define AppName             "ThermopacStructuringAgent"
#define AppPublisher        "Thermopac"
#define DesktopShortcutName "SolidWorks Structuring Agent"
#define AppURL              "https://thermopac-communication-thermopacllp.replit.app"
#define AppExeName          "run.bat"

; SourcePath = directory containing this .iss file (installer\)
; AgentRoot  = repo root (one level up from installer\)
; SourceDir  = dist\ThermopacStructuringAgent\  (built by build-installer.bat)
; PythonDir  = dist\python\
#define AgentRoot  SourcePath + ".."
#define SourceDir  AgentRoot + "\dist\ThermopacStructuringAgent"
#define PythonDir  AgentRoot + "\dist\python"

[Setup]
; New GUID — must differ from the Extraction Agent's GUID
AppId={{A7C4D2F1-9B3E-4E8A-B5F2-31D6C08EF450}
AppName={#AppName}
AppVersion={#MyAppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
; Install alongside (not replacing) the Extraction Agent
DefaultDirName={autopf}\ThermopacStructuringAgent
DefaultGroupName={#AppName}
AllowNoIcons=yes
; OutputDir is relative to this .iss file location (installer\)
; ..\installer_output resolves to installer_output\ at repo root
OutputDir=..\installer_output
OutputBaseFilename=ThermopacStructuringAgent-Setup-v{#MyAppVersion}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\run.bat
UninstallDisplayName={#AppName} v{#MyAppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";     Description: "Create a &desktop shortcut";                GroupDescription: "Additional icons:"
Name: "startmenuicon";   Description: "Create a &Start Menu shortcut";             GroupDescription: "Additional icons:"
Name: "startupschedule"; Description: "Start agent automatically at &Windows login"; GroupDescription: "Auto-start:"

[Files]
; ── Agent Python source (agent/ extractor/ structurer/ from dist\ThermopacStructuringAgent\)
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; ── Bundled Python runtime (optional — skipped if dist\python\ not present)
; When present the installer is fully self-contained.
; When absent, setup.ps1 downloads Python from python.org at install time.
Source: "{#PythonDir}\*"; DestDir: "{app}\python"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist

; ── Launch scripts (installer\ folder — not part of agent source)
Source: "{#SourcePath}run-service.bat";  DestDir: "{app}"; Flags: ignoreversion
Source: "{#SourcePath}makepy-repair.bat"; DestDir: "{app}"; Flags: ignoreversion

; ── PowerShell bootstrap (fetches Python at install time if not bundled)
Source: "{#SourcePath}setup.ps1"; DestDir: "{app}"; Flags: ignoreversion

; ── Default config — users edit after installation; never overwrite existing
Source: "{#AgentRoot}\config.ini"; DestDir: "{app}"; Flags: ignoreversion onlyifdoesntexist

; ── run.bat — placed at app root; built by build-installer.bat and listed here explicitly
Source: "{#AgentRoot}\run.bat"; DestDir: "{app}"; Flags: ignoreversion

; ── APPDATA auto-fix script — run.bat calls this on every startup
Source: "{#AgentRoot}\fix_appdata_url.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Staging root — writable by all users
Name: "{commonappdata}\ThermopacStructurer\temp"; Permissions: everyone-full
Name: "{commonappdata}\ThermopacStructurer\logs"; Permissions: everyone-full
; Staging output (default — users may override in config.ini)
Name: "C:\ThermopacStaging\drawings"; Permissions: everyone-full

[Icons]
; ── Start Menu group (always created)
Name: "{group}\{#AppName}";              Filename: "{app}\run.bat";            WorkingDir: "{app}"
Name: "{group}\Edit Config";             Filename: "notepad.exe";              Parameters: """{app}\config.ini"""
Name: "{group}\Repair COM Cache";        Filename: "{app}\makepy-repair.bat";  WorkingDir: "{app}"
Name: "{group}\Uninstall {#AppName}";    Filename: "{uninstallexe}"
; ── Named Start Menu shortcut (task-controlled)
Name: "{group}\{#DesktopShortcutName}";  Filename: "{app}\run.bat";           WorkingDir: "{app}"; Tasks: startmenuicon
; ── Desktop shortcut (task-controlled)
Name: "{autodesktop}\{#DesktopShortcutName}"; Filename: "{app}\run.bat";      WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\run.bat"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent unchecked shellexec

[Code]

// ── SolidWorks detection ──────────────────────────────────────────────────────
function DetectSwProgId(): String;
var
  Suffixes: array of String;
  i: Integer;
  Dummy: String;
begin
  Result := '';
  Suffixes := ['32','31','30','29','28','27'];
  for i := 0 to GetArrayLength(Suffixes) - 1 do
  begin
    if RegQueryStringValue(HKCR, 'SldWorks.Application.' + Suffixes[i], '', Dummy) then
    begin
      Result := 'SldWorks.Application.' + Suffixes[i];
      Exit;
    end;
  end;
end;

function CheckSolidWorks(): Boolean;
begin
  Result := DetectSwProgId() <> '';
end;

// ── makepy: pre-generate win32com COM type library cache ─────────────────────
procedure RunMakepy(ProgId: String);
var
  PyExe, Args: String;
  ResultCode: Integer;
begin
  PyExe := ExpandConstant('{app}\python\python.exe');
  if not FileExists(PyExe) then
  begin
    MsgBox(
      'Python not found at ' + PyExe + ' — skipping COM cache setup.' + #13#10 +
      'Run "Repair COM Cache" from the Start Menu after installation.',
      mbError, MB_OK);
    Exit;
  end;

  // Method A: standard makepy via ProgID
  Args := '-m win32com.client.makepy "' + ProgId + '"';
  Exec(PyExe, Args, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  if ResultCode = 0 then Exit;

  // Method B: gencache EnsureDispatch (last resort)
  Args := '-c "import win32com.client.gencache as g; g.EnsureDispatch(''' + ProgId + ''')"';
  Exec(PyExe, Args, '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  // Non-fatal — agent has its own 3-method DispatchEx fallback at runtime
end;

// ── Scheduled task ───────────────────────────────────────────────────────────
procedure CreateScheduledTask();
var
  ResultCode: Integer;
begin
  if IsTaskSelected('startupschedule') then
  begin
    Exec('schtasks.exe',
      '/Create /F /SC ONLOGON /RL HIGHEST /TN "ThermopacStructuringAgent" /TR "\"' +
      ExpandConstant('{app}') + '\run-service.bat\"" /DELAY 0000:30',
      '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode = 0 then
      MsgBox(
        'Scheduled task created: ThermopacStructuringAgent starts 30 seconds after login.',
        mbInformation, MB_OK)
    else
      MsgBox(
        'Warning: Could not create scheduled task (code ' + IntToStr(ResultCode) + '). ' +
        'Start from the Start Menu manually.', mbError, MB_OK);
  end;
end;

// ── Pre-install validation ───────────────────────────────────────────────────
function InitializeSetup(): Boolean;
begin
  Result := True;
  if not CheckSolidWorks() then
  begin
    MsgBox(
      'SolidWorks (2019–2024) was not detected on this machine.' + #13#10 + #13#10 +
      'ThermopacStructuringAgent requires SolidWorks to be installed and licensed.' + #13#10 +
      'Please install SolidWorks first, then re-run this installer.' + #13#10 + #13#10 +
      'Click OK to cancel.',
      mbCriticalError, MB_OK);
    Result := False;
  end;
end;

// ── Post-install ─────────────────────────────────────────────────────────────
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

    // If Python was NOT bundled, run setup.ps1 to download it
    if not FileExists(PyExe) then
    begin
      if FileExists(Ps1) then
      begin
        MsgBox(
          'Python was not bundled in this installer.' + #13#10 +
          'setup.ps1 will now download Python 3.11 from python.org.' + #13#10 + #13#10 +
          'A PowerShell window will open — leave it running until complete.' + #13#10 +
          'This requires internet access.',
          mbInformation, MB_OK);
        Exec('powershell.exe',
          '-ExecutionPolicy Bypass -File "' + Ps1 + '"',
          ExpandConstant('{app}'), SW_SHOW, ewWaitUntilTerminated, ResultCode);
      end else
        MsgBox(
          'Python not found and setup.ps1 is missing.' + #13#10 +
          'Re-run the installer or run bootstrap.bat manually.',
          mbError, MB_OK);
    end else
    begin
      // Python IS bundled — pre-generate makepy cache for detected SolidWorks
      ProgId := DetectSwProgId();
      if ProgId <> '' then
        RunMakepy(ProgId);
    end;

    // Force mode = testing in config.ini (works on fresh install AND upgrade)
    // SetIniString writes without BOM using Windows API — no manual edit needed
    SetIniString('agent', 'mode', 'testing', ExpandConstant('{app}\config.ini'));

    // ── Fix APPDATA api_url → dev server (no admin needed, user's APPDATA) ──
    // Creates the APPDATA dir if missing, then upserts the api_url key.
    // Preserves node_token and all other existing keys.
    ForceDirectories(ExpandConstant('{userappdata}\ThermopacStructuringAgent'));
    SetIniString('cloud', 'api_url',
      'https://5d05ae61-8225-4651-bb76-b4e20a4ddabb-00-3mex6zlihlmft.janeway.replit.dev',
      ExpandConstant('{userappdata}\ThermopacStructuringAgent\config.ini'));

    CreateScheduledTask();

    MsgBox(
      'Installation complete!' + #13#10 + #13#10 +
      'NEXT STEPS:' + #13#10 + #13#10 +
      '  1. Edit config.ini in the install folder:' + #13#10 +
      '       ' + ExpandConstant('{app}') + '\config.ini' + #13#10 + #13#10 +
      '     Required settings:' + #13#10 +
      '       [cloud]   node_id, node_token' + #13#10 +
      '       [structurer]  template_path  <-- MUST be set' + #13#10 + #13#10 +
      '  2. Run "ThermopacStructuringAgent" from the Start Menu or Desktop.' + #13#10 + #13#10 +
      '  3. If SolidWorks COM errors appear, run:' + #13#10 +
      '       Start Menu > ThermopacStructuringAgent > Repair COM Cache',
      mbInformation, MB_OK);
  end;
end;

// ── Uninstall cleanup ────────────────────────────────────────────────────────
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
begin
  if CurUninstallStep = usPostUninstall then
    Exec('schtasks.exe', '/Delete /F /TN "ThermopacStructuringAgent"',
         '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
