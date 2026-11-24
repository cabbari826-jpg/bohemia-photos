# Bohemia Photos — программа с окном

Окно: выбрать прайс (.xlsx), папку с PDF-каталогами и папку результата, нажать «Запустить».
Результат: папка `photos` (JPG по полному коду) и `report.xlsx` с миниатюрами и статусами.

## Как получить .exe
**Вариант 1, GitHub (ничего не ставить на компьютер).** Загрузить эту папку в репозиторий на GitHub.
Сборка запустится сама (Actions → Build Windows app), в результатах будут
`BohemiaPhotos.exe` (запускается сразу) и `BohemiaPhotos-Setup.exe` (установщик с ярлыком).

**Вариант 2, на своём Windows.** Поставить Python 3.12 с python.org (отметить «Add python.exe to PATH»),
дважды щёлкнуть `build_exe.bat`. Готовый файл: `dist\BohemiaPhotos.exe`.

## Запуск без сборки
    pip install pymupdf openpyxl pillow xlrd msoffcrypto-tool
    python bohemia_gui.py
