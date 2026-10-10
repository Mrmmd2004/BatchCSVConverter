; ===========================================================================
; Inno Setup script - Batch CSV Converter (x64 installer)
; Run build.bat first (it builds dist\BatchCsvConverter.exe), or build the exe
; manually and then compile this file with Inno Setup (F9).
; Output: release\BatchCsvConverter_Setup_x64.exe
; ===========================================================================

#define MyAppName "Batch CSV Converter"
#define MyAppVersion "2.1.0"
#define MyAppPublisher "Clubapp"
#define MyAppURL "https://github.com/Mrmmd2004/BatchCSVConverter"
#define MyAppExeName "BatchCsvConverter.exe"

[Setup]
AppId={{E1C6B6D2-6E2A-4B3E-9B7A-2C1F0B7B7A11}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}/issues
AppUpdatesURL={#MyAppURL}/releases
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=release
OutputBaseFilename=BatchCsvConverter_Setup_x64
SetupIconFile=clubapp.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
ArchitecturesAllowed=x64
PrivilegesRequiredOverridesAllowed=dialog
CloseApplications=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
Filename: "{#MyAppURL}"; Description: "View the project on GitHub"; Flags: shellexec nowait postinstall skipifsilent unchecked
