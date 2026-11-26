"""Обновление программы из релизов GitHub.

Установленная версия (через Inventa-Setup.exe) скачивает новый установщик и запускает его
в тихом режиме: он закрывает программу, ставит новую версию поверх и запускает её снова.
Версия без установки (просто Inventa.exe) открывает страницу релиза в браузере.
"""
import os
import re
import sys
import json
import tempfile
import subprocess
import urllib.request

REPO = 'cabbari826-jpg/bohemia-photos'
API = f'https://api.github.com/repos/{REPO}/releases/latest'
SETUP_NAME = 'Inventa-Setup.exe'
OLD_SETUP_NAME = 'BohemiaPhotos-Setup.exe'   # так установщик назывался до переименования


def current_version():
    """Версия из файла VERSION (в exe он вшит при сборке)."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    try:
        with open(os.path.join(base, 'VERSION'), encoding='utf-8') as f:
            return f.read().strip()
    except OSError:
        return '0.0.0'


def parse(v):
    return tuple(int(x) for x in re.findall(r'\d+', str(v))[:3]) or (0,)


def is_installed():
    """Программа поставлена установщиком (рядом лежит деинсталлятор Inno Setup)."""
    if not getattr(sys, 'frozen', False):
        return False
    folder = os.path.dirname(sys.executable)
    return any(f.lower().startswith('unins') and f.lower().endswith('.exe') for f in os.listdir(folder))


def latest(timeout=10):
    """-> dict(version, setup_url, page, notes) последнего релиза или None."""
    req = urllib.request.Request(API, headers={'Accept': 'application/vnd.github+json',
                                               'User-Agent': 'Inventa-updater'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        d = json.load(r)
    assets = {a.get('name'): a['browser_download_url'] for a in d.get('assets', [])}
    setup = assets.get(SETUP_NAME) or assets.get(OLD_SETUP_NAME)
    return dict(version=d.get('tag_name', '').lstrip('v'), setup_url=setup, page=d.get('html_url'),
                notes=d.get('body') or '')


def check():
    """-> сведения о релизе, если он новее текущей версии, иначе None. Ошибки сети пробрасываются."""
    rel = latest()
    if rel and rel['version'] and parse(rel['version']) > parse(current_version()):
        return rel
    return None


def download(url, progress=None):
    """Скачивает установщик во временную папку, возвращает путь."""
    dest = os.path.join(tempfile.gettempdir(), SETUP_NAME)
    req = urllib.request.Request(url, headers={'User-Agent': 'Inventa-updater'})
    with urllib.request.urlopen(req, timeout=60) as r, open(dest, 'wb') as f:
        total = int(r.headers.get('Content-Length') or 0)
        done = 0
        while True:
            chunk = r.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if progress and total:
                progress(done / total)
    return dest


def _state_dir():
    d = os.path.join(os.environ.get('LOCALAPPDATA') or tempfile.gettempdir(), 'Inventa')
    os.makedirs(d, exist_ok=True)
    return d


MARK = 'update.json'
LOG_NAME = 'Inventa-update.log'


def _writable(folder):
    try:
        t = os.path.join(folder, '.inventa-write-test')
        with open(t, 'w'):
            pass
        os.remove(t)
        return True
    except OSError:
        return False


def installer_args(folder, log, all_users=False):
    a = ['/SP-', '/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CLOSEAPPLICATIONS', '/NOCANCEL',
         f'/LOG={log}', f'/DIR={folder}']
    if all_users:
        a.append('/ALLUSERS')
    return a


def _quote(a):
    # /DIR=C:\Program Files\Inventa -> "/DIR=C:\Program Files\Inventa" (так понимает Inno Setup)
    return f'"{a}"' if ' ' in a else a


def run_installer(path, version=''):
    """Запускает установщик тихо в ту же папку, где стоит программа; он закроет программу,
    обновит и запустит её снова. Ход установки пишется в журнал %TEMP%\\Inventa-update.log;
    при следующем запуске программа проверит, что версия действительно сменилась (pending_result)."""
    folder = os.path.dirname(sys.executable)
    log = os.path.join(tempfile.gettempdir(), LOG_NAME)
    # программа стоит в Program Files (ставили «для всех») — без прав администратора туда не записать
    admin = not _writable(folder)
    args = installer_args(folder, log, admin)
    try:
        with open(os.path.join(_state_dir(), MARK), 'w', encoding='utf-8') as f:
            json.dump(dict(target=version, old=current_version(), log=log, folder=folder), f)
    except OSError:
        pass
    if sys.platform.startswith('win'):
        import ctypes
        # как двойной щелчок по файлу: отдельный процесс, не связанный с программой; для Program Files — запрос UAC
        rc = ctypes.windll.shell32.ShellExecuteW(None, 'runas' if admin else 'open', path,
                                                 ' '.join(_quote(a) for a in args), None, 1)
        if rc > 32:
            return
        if admin:
            raise OSError('нужны права администратора: программа установлена в ' + folder)
    subprocess.Popen([path] + args, close_fds=True)


def pending_result():
    """После обновления: None — проверять нечего; ('ok', версия) или ('fail', версия, хвост журнала)."""
    p = os.path.join(_state_dir(), MARK)
    try:
        with open(p, encoding='utf-8') as f:
            d = json.load(f)
    except (OSError, ValueError):
        return None
    try:
        os.remove(p)
    except OSError:
        pass
    target = d.get('target') or ''
    if target and parse(current_version()) >= parse(target):
        return ('ok', current_version())
    tail = ''
    try:
        with open(d.get('log') or '', encoding='utf-8', errors='replace') as f:
            lines = [ln.rstrip() for ln in f if ln.strip()]
        tail = '\n'.join(lines[-12:])
    except OSError:
        tail = 'журнал установки не найден — установщик, похоже, не запустился (антивирус?)'
    return ('fail', target, tail)
