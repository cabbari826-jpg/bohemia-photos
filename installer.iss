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
; обновление программы, поставленной «для всех» в Program Files, запускается с /ALLUSERS
PrivilegesRequiredOverridesAllowed=commandline
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
// Программа сама закрывается после запуска обновления — ждём (до 20 с), пока её exe освободится
function WaitUnlocked(f: String): Boolean;
var
  i: Integer;
begin
  Result := True;
  for i := 1 to 40 do
  begin
    if not FileExists(f) then Exit;
    if RenameFile(f, f + '.upd') then
    begin
      RenameFile(f + '.upd', f);
      Exit;
    end;
    Sleep(500);
  end;
  Result := False;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  old: String;
begin
  Result := '';
  old := ExpandConstant('{app}\BohemiaPhotos.exe');
  if WaitUnlocked(old) then
    DeleteFile(old);
  if not WaitUnlocked(ExpandConstant('{app}\Inventa.exe')) then
    Log('Inventa.exe всё ещё занят — пробуем закрыть через Restart Manager');
end;
