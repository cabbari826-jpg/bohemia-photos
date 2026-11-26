@echo off
chcp 65001 >nul
REM Сборка BohemiaPhotos.exe. Нужен Python 3.10+ с python.org (галочка "Add to PATH").
cd /d "%~dp0"
python -m pip install -r requirements.txt || goto :err
python make_version_info.py || goto :err
python -m PyInstaller --noconfirm --onefile --windowed --name BohemiaPhotos --add-data "VERSION;." --version-file version_info.txt bohemia_gui.py || goto :err
echo.
echo Готово: dist\BohemiaPhotos.exe
pause
exit /b 0
:err
echo Ошибка сборки
pause
exit /b 1
