; MailWeave — Inno Setup installer script
; Compile with:  ISCC.exe installer.iss
; Output:        installer_output\MailWeave-Setup-1.0.exe

#define AppName      "MailWeave"
; AppVersion is pulled from _version.iss, which sync_version.py regenerates
; from version.py before each build.
#include "_version.iss"
#define AppPublisher "MailWeave Professional"
#define AppExeName   "MailWeave.exe"
#define AppDesc      "Bundle Outlook emails into PDF or Word documents"
#define SourceDir    "dist\MailWeave"

[Setup]
AppId={{E7F3B2A1-4C8D-4F9E-B123-5A6D7E8F9012}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL=
AppSupportURL=
AppUpdatesURL=
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
PrivilegesRequired=admin
OutputDir=installer_output
OutputBaseFilename=MailWeave-Setup-{#AppVersion}
SetupIconFile=mailweave.ico
WizardStyle=modern
WizardImageFile=wizard_side.bmp
WizardSmallImageFile=wizard_top.bmp
Compression=lzma2/ultra64
SolidCompression=yes
; Minimum Windows 10
MinVersion=10.0
ArchitecturesInstallIn64BitMode=x64compatible
; Uninstaller branding
UninstallDisplayIcon={app}\{#AppExeName}
UninstallDisplayName={#AppName} {#AppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Bundle the entire PyInstaller one-dir output
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Start Menu
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Comment: "{#AppDesc}"
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"
; Desktop (only if task selected)
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Comment: "{#AppDesc}"; Tasks: desktopicon

[Run]
; Offer to launch the app after install
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up settings written to %APPDATA%\MailWeave on uninstall (optional — comment out to keep user settings)
; Type: filesandordirs; Name: "{userappdata}\MailWeave"
