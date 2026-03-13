; ═══════════════════════════════════════════════════════════
;  Equipment Maintenance Manager — Inno Setup Script
;  by Clan  |  v1.0.0
; ═══════════════════════════════════════════════════════════

#define AppName      "Maintenance Manager"
#define AppVersion   "1.0.0"
#define AppPublisher "Clan"
#define AppExeName   "Maintenance Manager.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL=
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppPublisher}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
LicenseFile=LICENSE.txt
InfoAfterFile=README.txt
OutputDir=installer_output
OutputBaseFilename=Maintenance Manager Setup
SetupIconFile=icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon";    Description: "Create a &desktop shortcut";    GroupDescription: "Additional icons:"; Flags: unchecked
Name: "startmenuicon";  Description: "Create a &Start Menu shortcut"; GroupDescription: "Additional icons:"; Flags: checkedonce

[Files]
; Main executable — built by PyInstaller in the dist\ folder
Source: "dist\Maintenance Manager.exe"; DestDir: "{app}"; Flags: ignoreversion

; Docs
Source: "LICENSE.txt"; DestDir: "{app}"; Flags: ignoreversion
Source: "README.txt";  DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"
Name: "{group}\README";               Filename: "{app}\README.txt"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"

; Desktop (optional task)
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; \
  Description: "Launch {#AppName}"; \
  Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Remove files the app creates at runtime so the folder is clean on uninstall
Type: files;     Name: "{app}\equipment.db"
Type: files;     Name: "{app}\config.json"
Type: files;     Name: "{app}\maintenance_manager.log"
Type: filesandordirs; Name: "{app}"
