; Inno Setup script for Specification Formatter (Windows desktop app).
;
; Compiled by .github/workflows/release.yml with:
;   ISCC /DMyAppVersion=1.2.3 packaging\windows\installer.iss
; and expects the PyInstaller one-folder output at dist\SpecificationFormatter\.
;
; Produces dist\installer\SpecificationFormatterSetup.exe -- a normal
; double-click installer with a Start-menu shortcut, an optional desktop icon,
; and a clean uninstaller. The app is NOT code-signed, so Windows SmartScreen
; shows a "Windows protected your PC" notice on first run (More info ->
; Run anyway); that is expected and documented in docs/RELEASE_WINDOWS.md and
; the README.

#ifndef MyAppVersion
  #define MyAppVersion "0.0.0"
#endif

#define MyAppName "Specification Formatter"
#define MyAppPublisher "Abraham Borg"
#define MyAppExeName "SpecificationFormatter.exe"
#define MyAppURL "https://github.com/abe-borg/spec-template-normalizer"

[Setup]
; A stable AppId ties every version together so an install upgrades in place
; instead of stacking side-by-side. Do NOT change this GUID across releases.
; This GUID is unique to Specification Formatter (distinct from any sibling app)
; so the two never collide in Add/Remove Programs or on disk.
AppId={{BA534C27-0FC9-4DE9-B307-DBCFB3256A49}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases/latest
DefaultDirName={autopf}\Specification Formatter
DefaultGroupName=Specification Formatter
DisableProgramGroupPage=yes
; Per-user install: no admin/UAC prompt, which keeps the unsigned experience as
; smooth as possible (the user only sees the one SmartScreen notice, not an
; elevation prompt on top of it).
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\..\dist\installer
OutputBaseFilename=SpecificationFormatterSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
; Let an in-place update replace the running app: Inno detects a running
; instance and offers to close it. Pairs with the in-app updater, which exits
; the app before launching this installer.
CloseApplications=yes
RestartApplications=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; The entire PyInstaller one-folder output.
Source: "..\..\dist\SpecificationFormatter\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
