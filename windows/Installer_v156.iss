#define MyAppVersion "15.6.1"
#define MyAppExe "DailyWorkReports.exe"

[Setup]
; Matches the historical ProTech user-scoped release (v15.3.1).
; Inno Setup uses this stable ID to upgrade in place without uninstalling old data.
AppId={{D5874C34-31C8-4E48-AE9B-EC4D2CC51100}
AppName=Daily Work Reports
AppVersion={#MyAppVersion}
AppVerName=Daily Work Reports {#MyAppVersion}
AppPublisher=ProTech
AppPublisherURL=https://alghad.digital/
AppSupportURL=https://alghad.digital/programs/reports/
DefaultDirName={localappdata}\Programs\ProTech\Daily Work Reports
DefaultGroupName=ProTech
DisableProgramGroupPage=yes
OutputDir=E:\ProTech\Releases\DailyWorkReports_Update_v15.6.1
OutputBaseFilename=DailyWorkReports_Update_v15.6.1
PrivilegesRequired=lowest
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
CloseApplications=yes
RestartApplications=no
UsePreviousAppDir=yes
DirExistsWarning=no
; All report files stay in LOCALAPPDATA outside {app}; [UninstallDelete]
; is deliberately omitted so upgrades and uninstall never erase user data.
UninstallDisplayIcon={app}\{#MyAppExe}
ArchitecturesAllowed=x64compatible

[Files]
Source: "E:\ProTech\Releases\DailyWorkReports_Update_v15.6.1\DailyWorkReports_Public_v15.6.1.exe"; DestDir: "{app}"; DestName: "{#MyAppExe}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Daily Work Reports"; Filename: "{app}\{#MyAppExe}"
Name: "{autodesktop}\Daily Work Reports"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "إنشاء اختصار على سطح المكتب"; GroupDescription: "اختصارات إضافية:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "تشغيل Daily Work Reports"; Flags: nowait postinstall skipifsilent
