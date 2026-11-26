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


def run_installer(path):
    """Запускает установщик тихо; он сам закроет программу, обновит и запустит её снова."""
    flags = 0x00000008 | 0x00000200 if sys.platform.startswith('win') else 0   # DETACHED | NEW_PROCESS_GROUP
    subprocess.Popen([path, '/SILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/CLOSEAPPLICATIONS'],
                     creationflags=flags, close_fds=True)
