; Inno Setup: установщик для BohemiaPhotos.exe (собирается после PyInstaller)
#define AppVer GetEnv("APP_VERSION")
#if AppVer == ""
  #define AppVer "1.0.0"
#endif
[Setup]
AppName=Bohemia Photos
AppVersion={#AppVer}
DefaultDirName={autopf}\Bohemia Photos
DefaultGroupName=Bohemia Photos
OutputDir=dist
OutputBaseFilename=BohemiaPhotos-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "ru"; MessagesFile: "compiler:Languages\Russian.isl"

[Tasks]
Name: "desktopicon"; Description: "Ярлык на рабочем столе"

[Files]
Source: "dist\BohemiaPhotos.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Bohemia Photos"; Filename: "{app}\BohemiaPhotos.exe"
Name: "{autodesktop}\Bohemia Photos"; Filename: "{app}\BohemiaPhotos.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\BohemiaPhotos.exe"; Description: "Запустить"; Flags: nowait postinstall
