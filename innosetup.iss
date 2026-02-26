#define MyAppName      "OneNoteBackupExporter"
#define MyAppVersion   "1.0.1"
#define MyAppPublisher "JLI Software"
#define MyAppURL       "https://github.com/vikingjunior12/OneNoteBackupExporter"
#define MyAppExeName   "OneNoteExporter.exe"
#define MyBuildDir     "C:\Users\JLi1\Nextcloud\C#_pwsh\OneNoteExporterC#\build"

[Setup]
AppId={{BBAE42F6-2F7E-474E-9A2D-C771DB475E3F}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
OutputDir=C:\Users\JLi1\Nextcloud\C#_pwsh\OneNoteExporterC#\
OutputBaseFilename=OneNoteBackupExporter_Setup_{#MyAppVersion}
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Gesamter Build-Ordner (self-contained, x86)
Source: "{#MyBuildDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}";  Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
