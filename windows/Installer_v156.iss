#define MyAppVersion "15.6.0"
#define MyAppExe "DailyWorkReports.exe"

[Setup]
AppId={{C2B01BDF-07F8-496A-AF48-BE9526F1269F}
AppName=Daily Work Reports
AppVersion={#MyAppVersion}
AppVerName=Daily Work Reports {#MyAppVersion}
AppPublisher=ProTech
AppPublisherURL=https://alghad.digital/
AppSupportURL=https://alghad.digital/programs/reports/
DefaultDirName={localappdata}\Programs\ProTech\Daily Work Reports
DefaultGroupName=ProTech
DisableProgramGroupPage=yes
OutputDir=E:\ProTech\Releases\DailyWorkReports_Public_v15.6.0
OutputBaseFilename=Setup_DailyWorkReports_v15.6.0
PrivilegesRequired=lowest
WizardStyle=modern
Compression=lzma2
SolidCompression=yes
CloseApplications=yes
RestartApplications=yes
UsePreviousAppDir=yes
UninstallDisplayIcon={app}\{#MyAppExe}
ArchitecturesAllowed=x64compatible

[Files]
Source: "E:\ProTech\Releases\DailyWorkReports_Public_v15.6.0\DailyWorkReports_Public_v15.6.0.exe"; DestDir: "{app}"; DestName: "{#MyAppExe}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Daily Work Reports"; Filename: "{app}\{#MyAppExe}"
Name: "{autodesktop}\Daily Work Reports"; Filename: "{app}\{#MyAppExe}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "إنشاء اختصار على سطح المكتب"; GroupDescription: "اختصارات إضافية:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExe}"; Description: "تشغيل Daily Work Reports"; Flags: nowait postinstall skipifsilent
