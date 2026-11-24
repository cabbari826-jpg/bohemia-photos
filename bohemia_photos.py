#!/usr/bin/env python3
"""
Bohemia: фото товаров из PDF-каталогов по кодам из прайс-листа.

Запуск:
    python bohemia_photos.py ПРАЙС.xlsx ПАПКА_С_КАТАЛОГАМИ [ПАПКА_РЕЗУЛЬТАТА] [ПАПКА_ИМЕЮЩИХСЯ_ФОТО]

Что делает:
  1. Читает все PDF в папке каталогов, находит короткие коды вида 1SI97/00000/630
     и фото товара над каждым кодом.
  2. Для каждой строки прайса сводит полный код 91L/1SI97/0/00000/630-660
     к короткому 1SI97/00000/630 и ищет его в каталогах.
  3. Сверяет описание из прайса (WHITE WINE 380 ml) с подписью в каталоге.
  4. Сохраняет фото: <результат>/photos/<полный код, / заменён на ->.jpg
     и отчёт <результат>/report.xlsx с миниатюрами и статусами.

Нужно: pip install pymupdf openpyxl pillow xlrd msoffcrypto-tool xlsxwriter
"""
import io
import os
import re
import sys
import glob
import collections

import pymupdf
import openpyxl
import xlsxwriter
from PIL import Image

CODE_RE = re.compile(r'^([0-9A-Z]{5})/([0-9A-Z]{5})/([0-9A-Z]{3})$')
DPI = 300

# Чем раньше в списке, тем выше приоритет каталога (более свежие первыми).
CATALOG_PRIORITY = ['news-2026', 'unity-2026', 'horeca-2025', 'crystal-bohemia-horeca-2025',
                    'news-2025', 'unity-2025', 'catalogue-unity', 'unity-2024', 'horeca-2024']


def catalog_rank(name):
    n = name.lower()
    for i, key in enumerate(CATALOG_PRIORITY):
        if key in n:
            return i
    m = re.search(r'20(\d\d)', n)
    return len(CATALOG_PRIORITY) + (100 - int(m.group(1)) if m else 100)


# ---------------------------------------------------------------- каталоги
def page_items(page):
    W, H = page.rect.width, page.rect.height
    imgs = []
    for info in page.get_image_info():
        r = pymupdf.Rect(info['bbox']) & page.rect
        if r.is_empty or r.width < 8 or r.height < 8:
            continue
        if r.width / r.height > 4:            # штрихкоды
            continue
        if r.width * r.height > 0.2 * W * H:  # фоновые и интерьерные фото
            continue
        imgs.append(r)
    words = page.get_text('words')
    codes = [(w[4].strip(), pymupdf.Rect(w[:4])) for w in words if CODE_RE.match(w[4].strip())]
    out = []
    for code, cr in codes:
        cx = (cr.x0 + cr.x1) / 2
        lab = [w for w in words if cr.y0 - 16 < w[3] <= cr.y0 + 1 and w[2] > cr.x0 - 25 and w[0] < cr.x1 + 25]
        label = ' '.join(w[4] for w in sorted(lab, key=lambda w: (round(w[1]), w[0])))
        # строка под кодом: объём / размеры (нужна для сверки описания)
        below = [w for w in words if cr.y1 - 1 <= w[1] < cr.y1 + 14 and w[2] > cr.x0 - 15 and w[0] < cr.x1 + 15]
        details = ' '.join(w[4] for w in sorted(below, key=lambda w: (round(w[1]), w[0])))
        cands = [r for r in imgs if r.x0 - 6 <= cx <= r.x1 + 6 and r.y1 <= cr.y0 + 4 and cr.y0 - r.y1 < 160]
        if not cands:
            continue
        cands.sort(key=lambda r: cr.y0 - r.y1)
        best = pymupdf.Rect(cands[0])
        is_set = 'set' in label.lower() or code.startswith('99999/')
        if is_set:
            # набор: объединяем все фото над подписью набора (графин + бокал)
            same_row = [(r.x0 + r.x1) / 2 for c2, r in codes
                        if c2 != code and abs(r.y0 - cr.y0) < 5]
            left = max([x for x in same_row if x < cx], default=cx - 180)
            right = min([x for x in same_row if x > cx], default=cx + 180)
            lo, hi = (left + cx) / 2, (cx + right) / 2
            for r in imgs:
                mx = (r.x0 + r.x1) / 2
                if lo <= mx <= hi and abs(r.y1 - best.y1) < 40 and r.y1 <= cr.y0 + 4:
                    best |= r
        shared = sum(1 for c2, r2 in codes
                     if best.x0 - 6 <= (r2.x0 + r2.x1) / 2 <= best.x1 + 6 and abs(r2.y0 - cr.y0) < 5)
        out.append(dict(code=code, label=label, details=details, rect=tuple(best), shared=shared, is_set=is_set))
    return out


def build_index(folder, log=print, progress=None):
    index = collections.defaultdict(list)
    pdfs = sorted(glob.glob(os.path.join(folder, '*.pdf')) + glob.glob(os.path.join(folder, '*.PDF')))
    pdfs = sorted(set(pdfs))
    for k, f in enumerate(pdfs):
        name = os.path.basename(f)
        log(f'  {name}')
        doc = pymupdf.open(f)
        for pn, page in enumerate(doc):
            if progress:
                progress(0.5 * (k + pn / max(len(doc), 1)) / max(len(pdfs), 1))
            for it in page_items(page):
                it.update(file=f, catalog=name, page=pn + 1)
                index[it['code']].append(it)
    return index, len(pdfs)


# ---------------------------------------------------------------- прайс
def short_code(full):
    """91L/1SI97/0/00000/630-660 -> 1SI97/00000/630"""
    parts = str(full).strip().split('/')
    if len(parts) < 5:
        return None
    return f'{parts[1]}/{parts[3]}/{parts[4][:3]}'


def find_columns(rows):
    """rows: список строк (кортежей) прайса, начиная с 1-й строки листа."""
    for i, row in enumerate(rows[:20], 1):
        heads = {}
        for col, v in enumerate(row, 1):
            if v not in (None, ''):
                heads.setdefault(str(v).strip().upper(), col)  # первая колонка с таким заголовком
        code = heads.get('CODE')
        desc = next((v for k, v in heads.items() if k.startswith('DESCRIPTION')), None)
        if code and desc:
            coll = next((v for k, v in heads.items() if k.startswith('COLLECTION')), None)
            return i, code, desc, coll
    # запасной вариант: раскладка SBS 2026
    return 1, 6, 11, 3


class PriceError(RuntimeError):
    """Понятная пользователю ошибка чтения прайса."""


def _xlsx_rows(src):
    try:
        wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    except Exception as e:
        raise PriceError('Не удалось открыть прайс как .xlsx. Откройте его в Excel и '
                         'сохраните как «Книга Excel (*.xlsx)».') from e
    return [tuple(r) for r in wb.worksheets[0].iter_rows(values_only=True)]


def read_price(path):
    """Возвращает (строки, номера строк с фото). Понимает .xlsx, .xls, защищённый .xlsx без пароля, .csv."""
    name = os.path.basename(path)
    if name.startswith('~$'):
        raise PriceError(f'«{name}» — временный файл Excel. Выберите сам прайс (без ~$ в начале имени).')
    try:
        with open(path, 'rb') as f:
            head = f.read(8)
    except PermissionError as e:
        raise PriceError('Нет доступа к файлу прайса. Закройте его в Excel и попробуйте снова.') from e
    if head.startswith(b'PK'):
        if name.lower().endswith('.xlsb'):
            raise PriceError('Формат .xlsb не поддерживается. Сохраните прайс как .xlsx.')
        rows = _xlsx_rows(path)
        try:
            had = photo_rows(path)
        except Exception:
            had = set()
        return rows, had
    if head.startswith(b'\xd0\xcf\x11\xe0'):
        # контейнер OLE: либо старый .xls, либо зашифрованный .xlsx
        import msoffcrypto
        with open(path, 'rb') as f:
            office = msoffcrypto.OfficeFile(f)
            encrypted = office.is_encrypted()
            if encrypted:
                buf = io.BytesIO()
                try:
                    office.load_key(password='VelvetSweatshop')  # пароль Excel «только для чтения»
                    office.decrypt(buf)
                except Exception as e:
                    raise PriceError('Прайс защищён паролем. Откройте его в Excel, снимите пароль '
                                     '(Файл → Сведения → Защита книги) и сохраните как .xlsx.') from e
                buf.seek(0)
                rows = _xlsx_rows(buf)
                buf.seek(0)
                try:
                    had = photo_rows(buf)
                except Exception:
                    had = set()
                return rows, had
        import xlrd
        try:
            book = xlrd.open_workbook(path)
        except Exception as e:
            raise PriceError('Не удалось прочитать старый формат .xls. Сохраните прайс как .xlsx.') from e
        sh = book.sheet_by_index(0)
        rows = [tuple(sh.row_values(i)) for i in range(sh.nrows)]
        return rows, set()  # фото в .xls не определяются
    # текстовый файл: CSV
    import csv
    for enc in ('utf-8-sig', 'cp1251'):
        try:
            with open(path, encoding=enc, newline='') as f:
                text = f.read()
            break
        except UnicodeDecodeError:
            continue
    else:
        raise PriceError('Не удалось распознать формат прайса. Сохраните его как .xlsx.')
    try:
        dialect = csv.Sniffer().sniff(text[:5000], delimiters=';,\t')
    except csv.Error:
        raise PriceError('Не удалось распознать формат прайса. Сохраните его как .xlsx.')
    rows = [tuple(r) for r in csv.reader(io.StringIO(text), dialect)]
    return rows, set()


def photo_rows(path):
    """Строки, в колонке A которых уже стоит фото."""
    rows = set()
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]
    for img in getattr(ws, '_images', []):
        try:
            rows.add(img.anchor._from.row + 1)
        except AttributeError:
            pass
    return rows


# ---------------------------------------------------------------- описание
SYN = {'whisky': 'set', 'drink': 'set', 'dof': 'tumbler', 'hb': 'tumbler', 'tof': 'tumbler',
       'old': 'tumbler', 'fashion': 'tumbler', 'highball': 'tumbler', 'champagne': 'flute'}
STOP = {'ml', 'cm', 'set', 'oz', 'fl'}


def tokens(s):
    s = s.lower().replace(',', '.')
    words = {SYN.get(w, w) for w in re.findall(r'[a-z]{2,}', s)}
    nums = set(re.findall(r'\d+(?:\.\d+)?', s))
    return words, nums


def nums_close(a, b):
    for x in a:
        for y in b:
            fx, fy = float(x), float(y)
            if fx and fy and abs(fx - fy) / max(fx, fy) <= 0.05:
                return True
    return False


def describe_match(desc, label, details='', code=''):
    if not desc:
        return 'нет описания', 0
    pw, pn = tokens(desc)
    lw, ln = tokens(label + ' ' + details)
    if code and 'ml' not in (label + ' ' + details).lower():
        # объём не напечатан (графины): берём из кода, 075 -> 750 ml
        v = code.split('/')[-1]
        if v.isdigit():
            ln |= {str(int(v) * 10), str(int(v))}
    pn -= {'1', '6'} if 'set' in pw else set()
    num_ok = not pn or nums_close(pn, ln)
    word_ok = bool((pw & lw) - STOP) or ('set' in pw and 'set' in lw)
    if num_ok and word_ok:
        return 'совпадает', 2
    if num_ok or word_ok:
        return 'проверить', 1
    return 'не совпадает', 0


# ---------------------------------------------------------------- вырезка
def crop(it, docs):
    doc = docs.get(it['file']) or docs.setdefault(it['file'], pymupdf.open(it['file']))
    page = doc[it['page'] - 1]
    r = pymupdf.Rect(it['rect'])
    r = pymupdf.Rect(r.x0 - 3, r.y0 - 3, r.x1 + 3, r.y1 + 3) & page.rect
    pix = page.get_pixmap(clip=r, dpi=DPI, alpha=False)
    return Image.frombytes('RGB', (pix.width, pix.height), pix.samples)


def safe_name(code):
    name = str(code).strip().replace('/', '-')
    return re.sub(r'[\\:*?"<>|\s]+', '_', name)


# ---------------------------------------------------------------- отчёт
def write_report(path, cols, widths, report, thumbs):
    """Отчёт с фото, вставленными в ячейки (как «Поместить в ячейку» в Excel):
    фото двигаются, сортируются и фильтруются вместе со строкой."""
    wb = xlsxwriter.Workbook(path)
    ws = wb.add_worksheet('Отчёт')
    base = dict(valign='vcenter', text_wrap=True)
    hdr = wb.add_format(dict(base, bold=True, bg_color='#F2F2F2', border=1))
    cell = wb.add_format(base)
    colored = {}
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)
    ws.write_row(0, 0, cols, hdr)
    photo_col = len(cols) - 1
    png = {}
    for r, (values, status_col, color, fname) in enumerate(report, 1):
        for c, v in enumerate(values):
            fmt = cell
            if c == status_col:
                fmt = colored.get(color) or colored.setdefault(color, wb.add_format(dict(base, bg_color=color)))
            ws.write(r, c, v, fmt)
        if fname:
            if fname not in png:
                buf = io.BytesIO()
                thumbs[fname].save(buf, 'JPEG', quality=85)
                png[fname] = buf.getvalue()
            ws.set_row(r, 90)
            ws.embed_image(r, photo_col, fname, {'image_data': io.BytesIO(png[fname]), 'description': fname})
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, len(report), len(cols) - 1)
    wb.close()


# ---------------------------------------------------------------- имеющиеся фото
IMG_EXT = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tif', '.tiff', '.gif')


_CYR_LOOKALIKE = str.maketrans('АВСЕНКМОРТХ', 'ABCEHKMOPTX')


def code_key(text):
    """Ключ для сравнения без учёта разделителей и регистра: 91L/1SI97/0/00000/630-660 -> 91L1SI97000000630660"""
    t = str(text).upper().translate(_CYR_LOOKALIKE)  # кириллица, похожая на латиницу
    return re.sub(r'[^0-9A-Z]', '', t)  # все разделители (/ - – — _ пробел . и т.д.) отбрасываются


def index_existing(folder, log=print):
    """Файлы фото в папке (с подпапками): ключ имени -> путь."""
    found = {}
    for root, _, files in os.walk(folder):
        for f in files:
            stem, ext = os.path.splitext(f)
            if ext.lower() in IMG_EXT and not f.startswith('~$'):
                found.setdefault(code_key(stem), os.path.join(root, f))
    log(f'  фото в папке: {len(found)}')
    return found


def find_existing(code, existing, keys_sorted):
    """Ищем фото для полного кода: точное совпадение, потом имя с припиской (…_1, «копия»),
    потом код без упаковочного хвоста (630-660 -> 630)."""
    if not existing:
        return None
    k = code_key(code)
    if k in existing:
        return existing[k]
    import bisect
    i = bisect.bisect_left(keys_sorted, k)
    if i < len(keys_sorted) and keys_sorted[i].startswith(k):
        return existing[keys_sorted[i]]
    parts = str(code).strip().split('/')
    if len(parts) >= 5:
        base = '/'.join(parts[:4] + [parts[4].split('-')[0]])
        kb = code_key(base)
        if kb in existing:
            return existing[kb]
    return None


def thumb_of(path):
    img = Image.open(path)
    img = img.convert('RGB')
    img.thumbnail((200, 200))
    return img


# ---------------------------------------------------------------- main
def run(price, folder, out, log=print, progress=None, existing_dir=None):
    """Основная работа. log(str) — вывод сообщений, progress(0..1) — прогресс.
    existing_dir — папка с уже имеющимися фото (необязательно): их берём в отчёт, из каталогов не режем."""
    os.makedirs(os.path.join(out, 'photos'), exist_ok=True)

    log('Читаю каталоги...')
    index, npdf = build_index(folder, log, progress)
    if not npdf:
        raise RuntimeError('В папке каталогов нет PDF-файлов')
    log(f'  найдено кодов в каталогах: {len(index)} (файлов: {npdf})')

    existing, keys_sorted = {}, []
    if existing_dir:
        log('Читаю папку с имеющимися фото...')
        existing = index_existing(existing_dir, log)
        keys_sorted = sorted(existing)

    log('Читаю прайс...')
    price_rows, had_photo = read_price(price)
    head, c_code, c_desc, c_coll = find_columns(price_rows)

    cols = ['Строка прайса', 'Коллекция', 'Код', 'Короткий код', 'Описание (прайс)',
            'Фото в прайсе', 'Статус', 'Подпись в каталоге', 'Проверка описания',
            'Каталог', 'Стр.', 'Файл фото', 'Фото']
    widths = [8, 22, 30, 18, 28, 9, 16, 30, 16, 34, 6, 34, 18]
    fills = {'есть в папке': '#DDEBF7', 'совпадает': '#C6EFCE', 'проверить': '#FFEB9C', 'не совпадает': '#FFC7CE',
             'нет в каталогах': '#FFC7CE'}
    report = []  # (значения, статус-колонка, цвет, имя фото)

    docs, saved, thumbs = {}, {}, {}
    stats = collections.Counter()
    all_rows = price_rows[head:]
    log('Вырезаю фото...')
    for rownum, row in enumerate(all_rows, head + 1):
        if progress and rownum % 20 == 0:
            progress(0.5 + 0.45 * (rownum - head) / max(len(all_rows), 1))
        code = row[c_code - 1] if len(row) >= c_code else None
        if not code or str(code).count('/') < 4 or not re.match(r'^[0-9A-Z]{3}/', str(code).strip()):
            continue
        code = str(code).strip()
        desc = str(row[c_desc - 1] or '').strip() if len(row) >= c_desc else ''
        coll = str(row[c_coll - 1] or '').strip() if c_coll and len(row) >= c_coll else ''
        sc = short_code(code)
        hits = index.get(sc, [])
        line = [rownum, coll, code, sc, desc, 'да' if rownum in had_photo else '']
        have = find_existing(code, existing, keys_sorted)
        if have:
            key = 'existing:' + have
            if key not in thumbs:
                try:
                    thumbs[key] = thumb_of(have)
                except Exception:
                    have = None
            if have:
                stats['есть в папке'] += 1
                rel = os.path.relpath(have, existing_dir)
                report.append((line + ['есть в папке', '', '', '', '', rel, ''], 6, fills['есть в папке'], key))
                continue
        if not hits:
            stats['нет в каталогах'] += 1
            report.append((line + ['нет в каталогах', '', '', '', '', '', ''], 6, fills['нет в каталогах'], None))
            continue
        scored = []
        for it in hits:
            verdict, score = describe_match(desc, it['label'], it.get('details', ''), it['code'])
            scored.append((-score, it['shared'] > 1, catalog_rank(it['catalog']), it, verdict))
        scored.sort(key=lambda x: x[:3])
        _, shared, _, best, verdict = scored[0]
        if shared and verdict == 'совпадает':
            verdict = 'проверить'
        stats[verdict] += 1
        fname = safe_name(code) + '.jpg'
        if fname not in saved:
            img = crop(best, docs)
            img.save(os.path.join(out, 'photos', fname), quality=92)
            th = img.copy()
            th.thumbnail((200, 200))
            saved[fname] = True
            thumbs[fname] = th
        report.append((line + ['найдено', (best['label'] + ' / ' + best['details']).strip(' /'), verdict,
                               best['catalog'], best['page'], fname, ''], 8, fills[verdict], fname))
    if not stats:
        raise PriceError('В прайсе не найдено ни одного кода вида 91L/1SI97/0/00000/630-660. '
                         'Проверьте, что выбран правильный файл и что в нём есть колонки CODE и DESCRIPTION.')
    log('Сохраняю отчёт...')
    write_report(os.path.join(out, 'report.xlsx'), cols, widths, report, thumbs)
    if progress:
        progress(1.0)

    total = sum(stats.values())
    log(f'\nАртикулов в прайсе: {total}')
    for k, v in stats.most_common():
        log(f'  {k}: {v}')
    log(f'Фото сохранено: {len(saved)} -> {os.path.join(out, "photos")}')
    log(f'Отчёт: {os.path.join(out, "report.xlsx")}')
    return stats, len(saved)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    price, folder = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else 'bohemia_result'
    existing = sys.argv[4] if len(sys.argv) > 4 else None
    run(price, folder, out, existing_dir=existing)


if __name__ == '__main__':
    main()
