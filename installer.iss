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
; всегда в папку пользователя (запись разрешена без прав администратора); старую копию удаляем сами
UsePreviousAppDir=no
OutputDir=dist
OutputBaseFilename=Inventa-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
; программу закрываем сами в PrepareToInstall (taskkill), окно Restart Manager «не удалось закрыть приложения» не нужно
CloseApplications=no
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
// Программа сама закрывается после запуска обновления — ждём (до 10 с), пока её exe освободится
function WaitUnlocked(f: String): Boolean;
var
  i: Integer;
  st: TFileStream;
begin
  // запущенный exe можно переименовать, но нельзя открыть на запись — проверяем именно запись
  Result := True;
  for i := 1 to 20 do
  begin
    if not FileExists(f) then Exit;
    try
      st := TFileStream.Create(f, fmOpenReadWrite or fmShareExclusive);
      st.Free;
      Exit;
    except
      Sleep(500);
    end;
  end;
  Result := False;
end;

procedure KillApp(name: String);
var
  rc: Integer;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /T /IM ' + name, '', SW_HIDE, ewWaitUntilTerminated, rc);
end;

// Удаляем прошлую установку (Bohemia Photos / Inventa) её же деинсталлятором, где бы она ни стояла.
// Для установки «для всех» (Program Files) Windows спросит разрешение; отказ — не ошибка, ставим дальше.
procedure UninstallOld(root: Integer);
var
  key, s: String;
  rc: Integer;
begin
  key := 'Software\Microsoft\Windows\CurrentVersion\Uninstall\Bohemia Photos_is1';
  if not RegQueryStringValue(root, key, 'UninstallString', s) then Exit;
  s := RemoveQuotes(s);
  if not FileExists(s) then Exit;
  Log('Удаляю прежнюю версию: ' + s);
  if not ShellExec('', s, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE, ewWaitUntilTerminated, rc) then
    Log('Деинсталлятор не запустился: ' + SysErrorMessage(rc))
  else
    Log('Деинсталлятор завершился, код ' + IntToStr(rc));
  Sleep(1000);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
begin
  Result := '';
  // программа сама закрывается после запуска обновления; ждём, потом закрываем принудительно
  if not WaitUnlocked(ExpandConstant('{app}\Inventa.exe')) then
    Log('Inventa.exe занят');
  KillApp('Inventa.exe');
  KillApp('BohemiaPhotos.exe');
  Sleep(500);
  UninstallOld(HKCU);
  UninstallOld(HKLM);
  if IsWin64 then
    UninstallOld(HKLM64);
  DeleteFile(ExpandConstant('{app}\BohemiaPhotos.exe'));
end;
