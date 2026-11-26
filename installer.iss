; Inno Setup: установщик для Inventa.exe (собирается после PyInstaller)
; Раньше программа называлась Bohemia Photos. AppId оставлен прежним, поэтому новая версия
; ставится поверх старой (та же запись в «Установке и удалении программ»), а старые exe и ярлыки удаляются.
#define AppVer GetEnv("APP_VERSION")
#if AppVer == ""
  #define AppVer "1.0.0"
#endif
[Setup]
AppId=Bohemia Photos
AppName=Inventa
AppVersion={#AppVer}
AppPublisher=Jooble Soft
AppCopyright=© Jooble Soft
VersionInfoCompany=Jooble Soft
VersionInfoProductName=Inventa
VersionInfoDescription=Установка Inventa
VersionInfoVersion={#AppVer}
SetupIconFile=app.ico
UninstallDisplayIcon={app}\Inventa.exe
UninstallDisplayName=Inventa (Jooble Soft)
DefaultDirName={autopf}\Inventa
DefaultGroupName=Inventa
UsePreviousGroup=no
OutputDir=dist
OutputBaseFilename=Inventa-Setup
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

[InstallDelete]
; остатки старого названия
Type: files; Name: "{app}\BohemiaPhotos.exe"
Type: files; Name: "{autodesktop}\Bohemia Photos.lnk"
Type: filesandordirs; Name: "{autoprograms}\Bohemia Photos"

[Files]
Source: "dist\Inventa.exe"; DestDir: "{app}"

[Icons]
Name: "{group}\Inventa"; Filename: "{app}\Inventa.exe"
Name: "{autodesktop}\Inventa"; Filename: "{app}\Inventa.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Inventa.exe"; Description: "Запустить"; Flags: nowait postinstall

[Code]
// Старая версия (BohemiaPhotos.exe) сама закрывается после запуска обновления — ждём, пока файл освободится
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  i: Integer;
  f: String;
begin
  Result := '';
  f := ExpandConstant('{app}\BohemiaPhotos.exe');
  for i := 1 to 30 do
  begin
    if (not FileExists(f)) or DeleteFile(f) then
      Break;
    Sleep(500);
  end;
end;
