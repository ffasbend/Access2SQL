#define MyAppName "Access2SQL"
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif
#define MyAppPublisher "ffasbend"
#define MyAppExeName "Access2SQL.exe"

[Setup]
AppId={{A2B3E0A1-4C9A-4A2D-B2F0-ACCESS2SQL}}
AppName={#MyAppName}
AppVersion={#AppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Access2SQL
DefaultGroupName=Access2SQL
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=Access2SQL-{#AppVersion}-Windows-x64-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
UninstallDisplayIcon={app}\Access2SQL.exe

[Files]
Source: "..\dist\Access2SQL\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Access2SQL"; Filename: "{app}\Access2SQL.exe"
Name: "{autodesktop}\Access2SQL"; Filename: "{app}\Access2SQL.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Run]
Filename: "{app}\Access2SQL.exe"; Description: "Launch Access2SQL"; Flags: nowait postinstall skipifsilent
