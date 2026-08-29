; Inno Setup script for PstToEmlConverter
; Self-contained .NET 8 WPF app (win-x64). Requires desktop Microsoft Outlook at runtime.
; Compile with: ISCC.exe PstToEmlConverter.iss

#define MyAppName "PstToEmlConverter"
#define MyAppVersion "1.0.1"
#define MyAppPublisher "kgounaris"
#define MyAppExeName "PstToEmlConverter.exe"

[Setup]
; AppId uniquely identifies this application. KEEP THIS GUID CONSTANT across all future
; versions so upgrades and the uninstaller resolve to the same installation.
AppId={{5563236A-0528-40C3-B8A1-990BC295E1B0}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
; Leave the "Select Destination Location" page ENABLED (do NOT set DisableDirPage).
; Require admin so the default install target is Program Files.
PrivilegesRequired=admin
; 64-bit only: refuse non-x64 hosts and install in native 64-bit mode.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Shown before files are installed: the Outlook requirement.
InfoBeforeFile=InfoBefore.txt
OutputDir=Output
OutputBaseFilename={#MyAppName}-{#MyAppVersion}-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Entire published folder, including all subfolders.
Source: "..\PstToEmlConverter\bin\Release\net8.0-windows\publish\win-x64\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent

; No [UninstallDelete] / uninstaller entries needed: Inno Setup creates and registers the
; uninstaller automatically (unins000.exe + Apps & Features entry) from the [Files] above.
