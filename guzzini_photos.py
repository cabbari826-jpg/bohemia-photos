#!/usr/bin/env python3
"""
Guzzini и Helios: фото товаров из PDF-каталогов и журналов по кодам из прайс-листа.
Каталоги-таблицы «код — EAN» (Helios) разбирает table_catalog.py.

Запуск:
    python guzzini_photos.py ПРАЙС.xlsx ПАПКА_С_КАТАЛОГАМИ [ПАПКА_РЕЗУЛЬТАТА] [ПАПКА_ИМЕЮЩИХСЯ_ФОТО]

Как устроены коды:
  В каталоге у товара базовый код из 6 цифр: 281976 или 1793.00 (точка не важна).
  Рядом напечатаны коды цветов: 26, 82, 132, 255 или 156, 206…
  В прайсе код = базовый + цвет: 281976 + 82 -> 28197682, 1793.00 + 156 -> 179300156.
  Иногда код в прайсе полностью совпадает с кодом в каталоге (03297076).

Что делает:
  1. В каждом PDF находит базовые коды, подпись товара (блок текста с кодом), фото рядом
     и, если есть, отдельные фото по цветам (цветной образец с номером цвета).
  2. Для каждой строки прайса (все листы) ищет базовый код по началу кода и цвет по остатку.
  3. Сохраняет фото: <результат>/photos/<код>.jpg и отчёт <результат>/report.xlsx.
"""
import os
import re
import sys
import glob
import math
import collections

import pymupdf
import openpyxl
from PIL import Image

import bohemia_photos as core
import table_catalog

BASE_RE = re.compile(r'^(\d{4})\.(\d{2})$|^(\d{6})$|^(\d{8,9})$|^(\d{6}[A-Z]{2}\d{2})$')
COLOR_RE = re.compile(r'^\d{2,3}$')
MAX_LINK = 350          # дальше этого (pt) фото к подписи не относим


def base_key(text):
    t = text.strip().strip('.,;:()')
    m = BASE_RE.match(t)
    if not m:
        return None
    if m.group(1):
        return m.group(1) + m.group(2)
    return next(g for g in m.groups()[2:] if g)


def rect_dist(a, b):
    dx = max(0, b[0] - a[2], a[0] - b[2])
    dy = max(0, b[1] - a[3], a[1] - b[3])
    return math.hypot(dx, dy)


def area(r):
    return max(0, r[2] - r[0]) * max(0, r[3] - r[1])


def union(a, b):
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def overlap_ratio(a, b):
    i = (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))
    ia = area(i) if i[2] > i[0] and i[3] > i[1] else 0
    return ia / max(1, min(area(a), area(b)))


# ---------------------------------------------------------------- страница
def page_images(page):
    """Картинки страницы: [(rect, xref)]. xref = None, если картинка склеена из нескольких."""
    W, H = page.rect.width, page.rect.height
    out = []
    for info in page.get_image_info(xrefs=True):
        r = pymupdf.Rect(info['bbox']) & page.rect
        if r.is_empty or r.width < 15 or r.height < 15:
            continue
        if r.width / r.height > 5 or r.height / r.width > 5:      # полосы, разделители
            continue
        if r.width * r.height > 0.35 * W * H:                     # фон, интерьер на весь разворот
            continue
        t = tuple(r)
        tr = info.get('transform') or (1, 0, 0, 1, 0, 0)
        plain = tr[1] == 0 and tr[2] == 0 and tr[0] > 0 and tr[3] > 0 and pymupdf.Rect(info['bbox']) in page.rect
        xref = info.get('xref') if plain else None
        # фото с тенью/маской часто лежит двумя картинками одна на другой — сливаем
        for k, (o, ox) in enumerate(out):
            if overlap_ratio(o, t) > 0.8 and 0.6 < area(o) / max(1, area(t)) < 1.6:
                out[k] = (union(o, t), None)
                break
        else:
            out.append((t, xref or None))
    return out


def color_labels(words):
    """Подписи цветов: «82», «156» или в две строки «AR» / «00» -> AR00."""
    letters = [w for w in words if re.fullmatch(r'[A-Z]{2,3}', w[4])]
    digits = [w for w in words if COLOR_RE.match(w[4])]
    used, out = set(), []
    for lw in letters:
        below = [d for d in digits if abs(d[0] - lw[0]) < 5 and 0 < d[1] - lw[1] < 14]
        if below:
            d = min(below, key=lambda d: d[1])
            used.add(id(d))
            out.append((lw[4] + d[4], (lw[0], lw[1], max(lw[2], d[2]), d[3])))
    for d in digits:
        if id(d) not in used:
            out.append((d[4], tuple(d[:4])))
    return out


def caption_block(words, cw):
    """Подпись товара: строки, выровненные по левому краю кода, без больших разрывов по вертикали."""
    x0 = cw[0]
    col = [w for w in words if abs(w[0] - x0) < 4 or (w[0] > x0 - 4 and w[0] < x0 + 320)]
    lines = collections.defaultdict(list)
    for w in col:
        lines[round((w[1] + w[3]) / 4)].append(w)    # шаг 2pt
    starts = {k: v for k, v in lines.items() if min(w[0] for w in v) > x0 - 4 and
              any(abs(w[0] - x0) < 4 for w in v)}
    keys = sorted(starts)
    me = round((cw[1] + cw[3]) / 4)
    if me not in starts:
        return tuple(cw[:4]), []
    i = keys.index(me)
    sel = [me]
    for step in (-1, 1):
        j, prev = i, me
        while 0 <= j + step < len(keys):
            k = keys[j + step]
            ya = max(w[3] for w in starts[prev]) if step == 1 else min(w[1] for w in starts[prev])
            yb = min(w[1] for w in starts[k]) if step == 1 else max(w[3] for w in starts[k])
            if abs(yb - ya) > 16:
                break
            sel.append(k)
            prev, j = k, j + step
    ws = []
    for k in sel:
        # в строке берём только слова этой колонки (без разрыва больше 40pt)
        row = sorted(starts[k], key=lambda w: w[0])
        keep = [row[0]]
        for w in row[1:]:
            if w[0] - keep[-1][2] > 40:
                break
            keep.append(w)
        ws += keep
    r = (min(w[0] for w in ws), min(w[1] for w in ws), max(w[2] for w in ws), max(w[3] for w in ws))
    return r, ws


def page_items(page):
    """Товары на странице: базовый код, подпись, фото товара и фото по цветам."""
    if table_catalog.is_table_page(page):
        # каталог с таблицей «код — EAN» (Helios)
        return table_catalog.page_items(page, page_images)
    words = [w for w in page.get_text('words') if w[4].strip()]
    imgs = page_images(page)
    codes = []
    seen = set()
    for w in words:
        k = base_key(w[4])
        if not k or (k, round(w[0]), round(w[1])) in seen:
            continue
        seen.add((k, round(w[0]), round(w[1])))
        block, bw = caption_block(words, w)
        text = ' '.join(x[4] for x in sorted(bw, key=lambda x: (round(x[1]), x[0])) if x is not w)
        # цвета списком в подписи: строка из одних 2–3-значных чисел («156 206 257»)
        rows = collections.defaultdict(list)
        for x in bw:
            rows[round(x[1])].append(x[4])
        listed = []
        for vals in rows.values():
            if vals and all(COLOR_RE.match(v) for v in vals):
                listed += [v for v in vals if v not in listed]
        codes.append(dict(key=k, raw=w[4], word=w, block=block, text=text, listed=listed))
    if not codes:
        return []

    # образцы цвета: подпись цвета прямо над маленькой картинкой (или в её левом верхнем углу)
    W, H = page.rect.width, page.rect.height
    swatch, swatch_label = {}, {}
    for col, lr in color_labels(words):
        best, bd = None, None
        for k, (r, _) in enumerate(imgs):
            if area(r) > 0.06 * W * H:
                continue
            cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
            dx, dy = abs(cx - (lr[0] + 18)), cy - lr[3]
            if dx < 45 and -12 < dy < 100 and r[1] > lr[1] - 90:
                d = dx + dy * 0.5
                if bd is None or d < bd:
                    best, bd = k, d
        if best is not None and best not in swatch:
            swatch[best] = col
            swatch_label[best] = lr

    # ряд образцов принадлежит той подписи, что ближе к его началу (левой подписи цвета)
    forced = {}
    rows = []
    for k, lr in sorted(swatch_label.items(), key=lambda kv: (round(kv[1][1]), kv[1][0])):
        for row in rows:
            last = swatch_label[row[-1]]
            step = lr[0] - last[0]
            if abs(last[1] - lr[1]) < 4 and 0 < step < 130:
                if len(row) > 1:      # шаг ряда постоянный: соседний ряд другого товара не цепляем
                    first_step = swatch_label[row[1]][0] - swatch_label[row[0]][0]
                    if abs(step - first_step) > 0.2 * first_step:
                        continue
                row.append(k)
                break
        else:
            rows.append([k])
    for row in rows:
        first = swatch_label[row[0]]
        def score(n):
            b = codes[n]['block']
            return abs(b[0] - first[0]) + max(0, b[1] - first[3], first[1] - b[3])
        n = min(range(len(codes)), key=score)
        if score(n) <= MAX_LINK:
            for k in row:
                forced[k] = n

    # каждую картинку отдаём ближайшей подписи
    owner = {}
    for k, (r, _) in enumerate(imgs):
        best, bd = None, None
        for n, c in enumerate(codes):
            d = rect_dist(r, c['block'])
            if bd is None or d < bd:
                best, bd = n, d
        if k in forced:
            owner[k] = (forced[k], 0)
        elif bd is not None and bd <= MAX_LINK:
            owner[k] = (best, bd)

    out = []
    for n, c in enumerate(codes):
        mine = [k for k, (o, _) in owner.items() if o == n]
        colors = {}
        for k in mine:
            if k in swatch:
                colors.setdefault(swatch[k], imgs[k])
        photos = [k for k in mine if k not in swatch]
        main = None
        if photos:
            big = max(photos, key=lambda k: area(imgs[k][0]))
            main = imgs[big]
            # ряд одинаковых фото (товар во всех цветах) — фото по порядку цветов из подписи
            br = imgs[big][0]
            row = sorted([k for k in photos if abs(imgs[k][0][1] - br[1]) < 25 and
                          0.6 < area(imgs[k][0]) / max(1, area(br)) < 1.6], key=lambda k: imgs[k][0][0])
            if len(row) > 1:
                u = imgs[row[0]][0]
                for k in row[1:]:
                    u = union(u, imgs[k][0])
                main = (u, None)
                if c['listed'] and len(row) == len(c['listed']):
                    for col, k in zip(c['listed'], row):
                        colors.setdefault(col, imgs[k])
        all_colors = list(dict.fromkeys(c['listed'] + list(colors)))
        if main is None and not colors:
            continue
        out.append(dict(code=c['key'], raw=c['raw'], label=c['text'][:300], rect=main, colors=colors,
                        all_colors=all_colors, shared=0))
    # одно и то же фото у нескольких кодов — пометим
    cnt = collections.Counter(tuple(round(v) for v in it['rect'][0]) for it in out if it['rect'])
    for it in out:
        if it['rect']:
            it['shared'] = cnt[tuple(round(v) for v in it['rect'][0])]
    return out


def catalog_rank(name):
    m = re.findall(r'20(\d\d)', name)
    year = int(m[-1]) if m else 0
    # каталоги точнее журналов: при равной оценке берём каталог
    return (-year, 1 if 'magazine' in name.lower() else 0, name)


def build_index(folder, log=print, progress=None):
    index = collections.defaultdict(list)
    pdfs = sorted(set(glob.glob(os.path.join(folder, '*.pdf')) + glob.glob(os.path.join(folder, '*.PDF'))))
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
ITEM_HEAD = re.compile(r'^(item|code|codice|cod\.?|article|articolo|артикул|код|item\s*code|ref\.?|item\s*no\.?|'
                       r'referenz\s*(no|nr)\.?|reference\s*no\.?|artikel\s*(no|nr)\.?)$', re.I)
DESC_HEAD = re.compile(r'desc|наимен|описан|\bname\b|designation|bezeichnung', re.I)


def read_price_all(path, log=print):
    """Все листы прайса: [(лист, № строки, код, описание, цвет, есть фото в прайсе)]."""
    name = os.path.basename(path)
    if name.lower().endswith('.pdf'):
        import invoice_reader
        sheets = [('', invoice_reader.read_pdf_price(path, log=lambda *a: None), set())]
    elif not name.lower().endswith(('.xlsx', '.xlsm')):
        rows, had = core.read_price(path)
        sheets = [('', rows, had)]
    else:
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as e:
            raise core.PriceError('Не удалось открыть прайс. Сохраните его как .xlsx и закройте в Excel.') from e
        sheets = []
        for ws in wb.worksheets:
            had = set()
            for img in getattr(ws, '_images', []):
                try:
                    had.add(img.anchor._from.row + 1)
                except AttributeError:
                    pass
            sheets.append((ws.title, [tuple(r) for r in ws.iter_rows(values_only=True)], had))
    out = []
    for title, rows, had in sheets:
        head = None
        for i, row in enumerate(rows[:25]):
            texts = [str(v).strip() if v is not None else '' for v in row]
            code = next((c for c, t in enumerate(texts) if ITEM_HEAD.match(t)), None)
            desc = next((c for c, t in enumerate(texts) if DESC_HEAD.search(t)), None)
            if code is not None and desc is not None:
                color = next((c for c, t in enumerate(texts) if re.match(r'^(colou?r|colore|цвет)', t, re.I)), None)
                head = (i, code, desc, color)
                break
        if head is None:
            log(f'  лист «{title}»: нет колонки Item / Code — пропускаю')
            continue
        i, cc, dc, colc = head
        n = 0
        for rn, row in enumerate(rows[i + 1:], i + 2):
            v = row[cc] if cc < len(row) else None
            if v in (None, ''):
                continue
            if isinstance(v, float) and v.is_integer():
                v = int(v)
            code = str(v).strip()
            if not re.match(r'^[A-Za-z]{1,4}\s+\d[\w./-]*$', code):     # «STO 0141» — код целиком
                code = code.split()[0]
            if not re.search(r'\d{4}', code):
                continue
            desc = str(row[dc] or '').strip() if dc is not None and dc < len(row) else ''
            color = str(row[colc] or '').strip() if colc is not None and colc < len(row) else ''
            out.append(dict(sheet=title, row=rn, code=code, desc=desc, color=color, had=rn in had))
            n += 1
        log(f'  лист «{title}»: кодов {n}')
    if not out:
        raise core.PriceError('В прайсе не найдено колонки с кодами (Item / Code / Codice / Артикул).')
    return out


def match(code, index):
    """-> (список (позиция каталога, цвет или ''), как найдено)."""
    k = core.code_key(code)
    if k in index:
        return [(it, '') for it in index[k]], 'точно'
    for n in range(min(len(k) - 1, 9), 5, -1):
        if k[:n] in index:
            return [(it, k[n:]) for it in index[k[:n]]], 'по началу кода'
    m = re.match(r'^\s*(\d{4})-(\w[\w-]*)\s*$', str(code))      # Helios: 2941-777 -> модель 2941
    if m and m.group(1) in index:
        return [(it, core.code_key(m.group(2))) for it in index[m.group(1)]], 'по коду модели'
    return [], ''


# ---------------------------------------------------------------- вырезка
def crop(file, page_no, photo, docs):
    """Фото товара. Одиночную картинку берём прямо из PDF (без подписей поверх, в родном разрешении,
    прозрачный фон -> белый, поля обрезаем); склеенную — рендерим областью страницы."""
    rect, xref = photo
    doc = docs.get(file) or docs.setdefault(file, pymupdf.open(file))
    if xref:
        try:
            pix = pymupdf.Pixmap(doc, xref)
            smask = doc.xref_get_key(xref, 'SMask')
            if smask[0] == 'xref' and not pix.alpha:
                mask = pymupdf.Pixmap(doc, int(smask[1].split()[0]))
                if mask.width == pix.width and mask.height == pix.height:
                    pix = pymupdf.Pixmap(pix, mask)
            if pix.colorspace and pix.colorspace.n != 3:
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            mode = 'RGBA' if pix.alpha else 'RGB'
            img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
            if mode == 'RGBA':
                bg = Image.new('RGB', img.size, 'white')
                bg.paste(img, mask=img.split()[3])
                img = bg
            # обрезаем белые/прозрачные поля
            gray = img.convert('L').point(lambda v: 255 if v < 245 else 0)
            box = gray.getbbox()
            if box:
                pad = max(4, int(0.03 * max(img.size)))
                box = (max(0, box[0] - pad), max(0, box[1] - pad),
                       min(img.width, box[2] + pad), min(img.height, box[3] + pad))
                img = img.crop(box)
            if min(img.size) >= 24:
                # мелкую картинку увеличиваем до размера, как при печати 300 dpi (как и раньше при рендере)
                want = (rect[2] - rect[0]) * core.DPI / 72
                if img.width < want * 0.9 and pix.width < want:
                    k = want / pix.width
                    img = img.resize((max(1, int(img.width * k)), max(1, int(img.height * k))), Image.LANCZOS)
                img.info['native'] = pix.width, pix.height
                return img
        except Exception:
            pass
    return core.crop(dict(file=file, page=page_no, rect=rect), docs)


# ---------------------------------------------------------------- бренд прайса
BOHEMIA_CODE = re.compile(r'^[0-9A-Z]{3}/[0-9A-Z]{4,5}/')


HELIOS_CODE = re.compile(r'^\d{4}-\w{3}$')


def detect_brand(path):
    """'bohemia', 'guzzini' или 'helios' по имени файла и кодам в прайсе
    (91L/1SI97/… — Bohemia, 2941-001 и штрихкоды 4006657… — Helios, остальное — Guzzini и др.)."""
    name = os.path.basename(path).lower()
    for key in ('bohemia', 'guzzini', 'helios'):
        if key in name:
            return key
    try:
        if name.endswith('.pdf'):
            doc = pymupdf.open(path)
            text = ' '.join(doc[i].get_text() for i in list(range(min(len(doc), 3))) + [len(doc) - 1]).lower()
            return 'helios' if 'helios' in text else 'guzzini'
        if name.endswith(('.xlsx', '.xlsm')):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            rows = [r for ws in wb.worksheets for r in ws.iter_rows(max_row=60, values_only=True)]
        else:
            rows = core.read_price(path)[0][:60]
    except Exception:
        return 'bohemia'
    helios = 0
    for row in rows:
        for v in row:
            t = str(v).strip() if v is not None else ''
            if BOHEMIA_CODE.match(t):
                return 'bohemia'
            if HELIOS_CODE.match(t) or t.startswith('4006657') or 'helios' in t.lower():
                helios += 1
    return 'helios' if helios >= 3 else 'guzzini'


# ---------------------------------------------------------------- main
_INDEX_CACHE = {}   # повторный запуск с той же папкой каталогов не читает PDF заново


def run(price, folder, out, log=print, progress=None, existing_dir=None):
    os.makedirs(os.path.join(out, 'photos'), exist_ok=True)
    log('Читаю каталоги...')
    sig = (os.path.abspath(folder), tuple(sorted((f, os.path.getmtime(f)) for f in
                                                 glob.glob(os.path.join(folder, '*.[pP][dD][fF]')))))
    if sig in _INDEX_CACHE:
        index, npdf = _INDEX_CACHE[sig]
        log('  (каталоги уже прочитаны)')
    else:
        index, npdf = build_index(folder, log, progress)
        _INDEX_CACHE.clear()
        _INDEX_CACHE[sig] = (index, npdf)
    if not npdf:
        raise RuntimeError('В папке каталогов нет PDF-файлов')
    log(f'  найдено товаров в каталогах: {len(index)} (файлов: {npdf})')

    existing, keys_sorted = {}, []
    if existing_dir:
        log('Читаю папку с имеющимися фото...')
        existing = core.index_existing(existing_dir, log)
        keys_sorted = sorted(existing)

    log('Читаю прайс...')
    items = read_price_all(price, log)

    cols = ['Лист', 'Строка прайса', 'Код', 'Описание (прайс)', 'Цвет (прайс)', 'Фото в прайсе', 'Статус',
            'Код в каталоге', 'Цвет', 'Подпись в каталоге', 'Проверка', 'Каталог', 'Стр.', 'Файл фото', 'Фото']
    widths = [14, 8, 14, 32, 14, 9, 16, 12, 7, 40, 18, 30, 6, 24, 18]
    fills = {'есть в папке': '#DDEBF7', 'совпадает': '#C6EFCE', 'фото цвета': '#C6EFCE',
             'проверить цвет': '#FFEB9C', 'проверить': '#FFEB9C', 'фото модели': '#E2EFDA', 'нет в каталогах': '#FFC7CE'}
    report, docs, saved, thumbs, native = [], {}, {}, {}, {}
    # сколько строк прайса приходится на один базовый код (разные цвета одного товара)
    per_base = collections.Counter()
    for row in items:
        hits, _ = match(row['code'], index)
        if hits:
            per_base[hits[0][0]['code']] += 1
    stats = collections.Counter()
    log('Вырезаю фото...')
    for n, row in enumerate(items):
        if progress and n % 20 == 0:
            progress(0.5 + 0.45 * n / max(len(items), 1))
        code = row['code']
        line = [row['sheet'], row['row'], code, row['desc'], row['color'], 'да' if row['had'] else '']
        have = core.find_existing(code, existing, keys_sorted) if existing else None
        if have:
            key = 'existing:' + have
            try:
                thumbs.setdefault(key, core.thumb_of(have))
                stats['есть в папке'] += 1
                rel = os.path.relpath(have, existing_dir)
                report.append((line + ['есть в папке', '', '', '', '', '', '', rel, ''], 6, fills['есть в папке'], key))
                continue
            except Exception:
                pass
        hits, how = match(code, index)
        if not hits:
            stats['нет в каталогах'] += 1
            report.append((line + ['нет в каталогах'] + [''] * 8, 6, fills['нет в каталогах'], None))
            continue
        scored = []
        for it, color in hits:
            if color and color in it['colors']:
                verdict, rect, score = 'фото цвета', it['colors'][color], 0
            elif not color or (it['all_colors'] == [color]) or \
                    (not it['all_colors'] and it['shared'] <= 1 and per_base[it['code']] <= 1):
                verdict, rect, score = 'совпадает', it['rect'] or next(iter(it['colors'].values())), 1
            else:
                verdict, rect, score = 'проверить цвет', it['rect'] or next(iter(it['colors'].values())), 2
            if verdict == 'совпадает' and it['shared'] > 1:
                verdict, score = 'проверить', 2
            scored.append((score, catalog_rank(it['catalog']), it['page'], it, verdict, rect, color))
            if it.get('verdict'):
                v = it['verdict'] if not color else 'проверить цвет'
                scored[-1] = ({'фото цвета': 0, 'совпадает': 1, 'фото модели': 1}.get(v, 2), catalog_rank(it['catalog']),
                              it['page'], it, v, it['rect'], '')
        scored.sort(key=lambda x: x[:3])
        _, _, _, best, verdict, rect, color = scored[0]
        if color and color not in best['all_colors'] and best['all_colors']:
            note = f'цвета {color} нет среди: ' + ', '.join(best['all_colors'])
        else:
            note = ''
        stats[verdict] += 1
        fname = core.safe_name(code) + '.jpg'
        if fname not in saved:
            img = crop(best['file'], best['page'], rect, docs)
            if 'native' in img.info:
                native[fname] = img.info['native']
            img.save(os.path.join(out, 'photos', fname), quality=92)
            th = img.copy()
            th.thumbnail((200, 200))
            saved[fname] = True
            thumbs[fname] = th
        if max(native.get(fname, (999, 999))) < 160:
            note = (note + '; ' if note else '') + 'в каталоге фото мелкое'
        check = verdict + (f' ({note})' if note else '')
        report.append((line + ['найдено', best['raw'], color, best['label'], check, best['catalog'], best['page'],
                               fname, ''], 10, fills[verdict], fname))
    log('Сохраняю отчёт...')
    core.write_report(os.path.join(out, 'report.xlsx'), cols, widths, report, thumbs)
    if progress:
        progress(1.0)
    log(f'\nАртикулов в прайсе: {sum(stats.values())}')
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
    out = sys.argv[3] if len(sys.argv) > 3 else 'guzzini_result'
    existing = sys.argv[4] if len(sys.argv) > 4 else None
    run(price, folder, out, existing_dir=existing)


if __name__ == '__main__':
    main()
