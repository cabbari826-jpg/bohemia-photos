"""Свойства exe (Windows → Свойства → Подробно): издатель Jooble Soft, версия из файла VERSION."""
import os

COMPANY = 'Jooble Soft'
PRODUCT = 'Bohemia Photos'

here = os.path.dirname(os.path.abspath(__file__))
ver = open(os.path.join(here, 'VERSION'), encoding='utf-8').read().strip()
nums = [int(x) for x in (ver.split('.') + ['0'] * 4)[:4]]
t = tuple(nums)
text = f'''VSVersionInfo(
  ffi=FixedFileInfo(filevers={t}, prodvers={t}, mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('041904B0', [
      StringStruct('CompanyName', '{COMPANY}'),
      StringStruct('FileDescription', '{PRODUCT}'),
      StringStruct('FileVersion', '{ver}'),
      StringStruct('InternalName', 'BohemiaPhotos'),
      StringStruct('LegalCopyright', '© {COMPANY}'),
      StringStruct('OriginalFilename', 'BohemiaPhotos.exe'),
      StringStruct('ProductName', '{PRODUCT}'),
      StringStruct('ProductVersion', '{ver}')])]),
    VarFileInfo([VarStruct('Translation', [0x0419, 1200])])
  ]
)
'''
with open(os.path.join(here, 'version_info.txt'), 'w', encoding='utf-8') as f:
    f.write(text)
print('version_info.txt:', ver, COMPANY)
