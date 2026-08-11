; Inno Setup script for hoso Privacy Cleaner.
; Build the app first (build_exe.ps1), then compile this with ISCC.exe
; (see build_installer.ps1, which does both steps in order).

#define MyAppName "hoso Privacy Cleaner"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "hosoya"
#define MyAppExeName "hosoPrivacyCleaner.exe"

[Setup]
AppId={{9275B1E9-735E-4CDC-8622-A93FB1C6C7FC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=..\dist_installer
OutputBaseFilename=hosoPrivacyCleaner_Setup_{#MyAppVersion}
SetupIconFile=..\assets\app_icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
; Per-user by default so no admin rights are required to install.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Tasks]
Name: "desktopicon"; Description: "デスクトップにアイコンを作成する"; GroupDescription: "追加のショートカット:"; Flags: unchecked

[Files]
Source: "..\dist\hosoPrivacyCleaner\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
