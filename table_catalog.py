"""Каталоги с таблицей артикулов (Helios и похожие): «цвет | объём | ITEM NO. | EAN».

На развороте каталога у модели таблица: в каждой строке код с цветом (2941-001) и штрихкод EAN из 13 цифр,
слева от первой строки каждого цвета — цветной квадратик-образец. Фото товара (вырезанные, на прозрачном фоне)
лежат рядом с таблицей, часто по фото на цвет.

Для каждого кода берём фото, цвет которого ближе всего к образцу цвета в таблице. Если цвет подобрать
не удалось — фото модели (самое крупное у таблицы) с пометкой «проверить цвет».
Для базового кода модели (2941 — так он записан в прайсе) — фото модели.
"""
import re
import collections

import pymupdf

import bohemia_photos as core

CODE_RE = re.compile(r'^(\d{4})(?:-(\w{1,4}(?:-\w{1,4})?))?$')
EAN_RE = re.compile(r'^\d{12,16}$')
MAX_LINK = 520          # дальше этого (pt) фото к таблице не относим


def _center(r):
    return (r[0] + r[2]) / 2, (r[1] + r[3]) / 2


def _dist(a, b):
    dx = max(b[0] - a[2], a[0] - b[2], 0)
    dy = max(b[1] - a[3], a[1] - b[3], 0)
    return (dx * dx + dy * dy) ** 0.5


def _overlap(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    if w <= 0 or h <= 0:
        return 0.0
    return w * h / max(1, min((a[2] - a[0]) * (a[3] - a[1]), (b[2] - b[0]) * (b[3] - b[1])))


def code_rows(words):
    """Строки таблицы: слово-код, справа на той же высоте — EAN."""
    eans = [w for w in words if EAN_RE.match(w[4])]
    rows = []
    for w in words:
        m = CODE_RE.match(w[4])
        if not m:
            continue
        yc = (w[1] + w[3]) / 2
        # «3334-xxx» — цвет на выбор, штрихкода нет
        if re.match(r'^x{2,4}$', m.group(2) or '', re.I) or \
                any(abs((e[1] + e[3]) / 2 - yc) < 3 and 0 < e[0] - w[2] < 120 for e in eans):
            rows.append(dict(word=w, raw=w[4], base=m.group(1), color=m.group(2) or '', yc=yc))
    return rows


def is_table_page(page):
    return len(code_rows(page.get_text('words'))) >= 2


def swatches(page):
    """Цветные образцы в таблицах: небольшие залитые прямоугольники ~2:1."""
    out = []
    for d in page.get_drawings():
        r, fill = d['rect'], d.get('fill')
        if not fill or d.get('items') and len(d['items']) > 2:
            continue
        if 12 <= r.width <= 40 and 5 <= r.height <= 14 and 1.6 < r.width / r.height < 4.5:
            out.append((tuple(r), tuple(int(c * 255) for c in fill[:3])))
    return out


def tables(rows):
    """Строки с одинаковым x кода, идущие подряд, — одна таблица."""
    groups = []
    for r in sorted(rows, key=lambda r: (round(r['word'][0] / 4), r['yc'])):
        g = groups[-1] if groups else None
        if g and abs(g[-1]['word'][0] - r['word'][0]) < 4 and 0 < r['yc'] - g[-1]['yc'] < 40:
            g.append(r)
        else:
            groups.append([r])
    return groups


def image_color(page, photo, cache):
    """-> (средний цвет предмета RGB, доля «цветных» пикселей) по уменьшенной картинке."""
    key = tuple(round(v) for v in photo[0])
    if key in cache:
        return cache[key]
    rect, xref = photo
    doc = page.parent
    px = None
    try:
        if xref:
            pix = pymupdf.Pixmap(doc, xref)
            smask = doc.xref_get_key(xref, 'SMask')
            if smask[0] == 'xref' and not pix.alpha:
                mask = pymupdf.Pixmap(doc, int(smask[1].split()[0]))
                if mask.width == pix.width and mask.height == pix.height:
                    pix = pymupdf.Pixmap(pix, mask)
            if pix.colorspace and pix.colorspace.n != 3:
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            px = pix
    except Exception:
        px = None
    if px is None:
        px = page.get_pixmap(clip=pymupdf.Rect(rect), dpi=30)
    from PIL import Image
    img = Image.frombytes('RGBA' if px.alpha else 'RGB', (px.width, px.height), px.samples)
    img.thumbnail((80, 80))
    pts = []
    data = list(img.getdata())
    for p in data:
        if len(p) == 4:
            if p[3] < 160:
                continue
            p = p[:3]
        elif min(p) > 238:          # белый фон картинки без прозрачности
            continue
        pts.append(p)
    res = (pts, len(pts) / max(1, len(data)))
    cache[key] = res
    return res


def median_color(pts):
    if not pts:
        return None
    return tuple(sorted(p[i] for p in pts)[len(pts) // 2] for i in range(3))


def color_cost(img_rgb, sw_rgb):
    """Насколько цвет фото далёк от образца. На фото светотень: чёрное выходит тёмно-серым, белое — серым,
    поэтому яркость образца приводим к «фотографической» и сравниваем её мягче, чем оттенок."""
    def opp(c):
        r, g, b = c
        return r - g, (r + g) / 2 - b, (r + g + b) / 3
    a1, b1, l1 = opp(img_rgb)
    a2, b2, l2 = opp(sw_rgb)
    l2 = 40 + 0.65 * l2
    return 2 * ((a1 - a2) ** 2 + (b1 - b2) ** 2) ** 0.5 + abs(l1 - l2)


def page_items(page, page_images):
    """Позиции страницы в формате guzzini_photos.page_items (+ поле verdict)."""
    words = [w for w in page.get_text('words') if w[4].strip()]
    rows = code_rows(words)
    if not rows:
        return []
    W = page.rect.width
    imgs = [p for p in page_images(page)
            if (p[0][2] - p[0][0]) >= 40 and (p[0][3] - p[0][1]) >= 40]
    sw = swatches(page)
    groups = tables(rows)
    boxes = []
    for g in groups:
        cx0, top = g[0]['word'][0], g[0]['word'][1]
        # левый край таблицы — заголовок «FARBE/COLOUR» над строками, иначе ~260 pt левее кода
        heads = [w for w in words if w[4].upper() in ('FARBE', 'COLOUR', 'COLOR') and w[2] <= cx0
                 and 0 < top - w[3] < 30 and cx0 - w[0] < 320]
        x0 = min(w[0] for w in heads) if heads else cx0 - 260
        boxes.append((x0, top, max(r['word'][2] for r in g) + 70, g[-1]['word'][3]))
    # фото -> ближайшая таблица (на развороте — предпочтительно на той же половине)
    owner = collections.defaultdict(list)
    for k, (r, _) in enumerate(imgs):
        best = None
        for t, b in enumerate(boxes):
            d = _dist(r, b)
            (ax, ay), (bx, by) = _center(r), _center(b)
            d += 0.05 * ((ax - bx) ** 2 + (ay - by) ** 2) ** 0.5    # при равенстве — к той, что ближе центром
            if W > 800 and (_center(r)[0] < W / 2) != (_center(b)[0] < W / 2):
                d += 250
            if best is None or d < best[0]:
                best = (d, t)
        if best and best[0] < MAX_LINK:
            owner[best[1]].append(k)
    cache = {}
    out = []
    for t, g in enumerate(groups):
        mine = owner.get(t, [])
        main = imgs[max(mine, key=lambda k: (imgs[k][0][2] - imgs[k][0][0]) * (imgs[k][0][3] - imgs[k][0][1]))] \
            if mine else None
        # цвет строки: ближайший образец выше/на уровне строки, левее кода
        cx = g[0]['word'][0]
        bx0 = boxes[t][0]
        tsw = [s for s in sw if bx0 - 5 <= s[0][0] and s[0][2] < cx - 20 and g[0]['yc'] - 12 <= (s[0][1] + s[0][3]) / 2 <= g[-1]['yc'] + 6]
        name_of = {}
        for r in g:
            above = [s for s in tsw if (s[0][1] + s[0][3]) / 2 <= r['yc'] + 4]
            s = max(above, key=lambda s: s[0][1]) if above else None
            r['swatch'] = s
            if s:
                nm = [w[4] for w in words if s[0][2] < w[0] < cx - 40 and abs((w[1] + w[3]) / 2 - (s[0][1] + s[0][3]) / 2) < 5]
                name_of[s[0]] = ' '.join(nm[:4])
        # подбор фото по цвету: каждому образцу — фото с наибольшей долей его цвета
        pick = {}
        if mine and tsw:
            cand = []
            for k in mine:
                r = imgs[k][0]
                # область страницы (не отдельная картинка), перекрытая другим фото, — цвет не определить
                if imgs[k][1] is None and any(j != k and _overlap(r, imgs[j][0]) > 0.25 for j in range(len(imgs))):
                    continue
                mc = median_color(image_color(page, imgs[k], cache)[0])
                if mc is None:
                    continue
                for s in tsw:
                    cand.append((color_cost(mc, s[1]), s[0], k))
            cand.sort()
            used_s, used_k = set(), set()
            for cost, s, k in cand:
                if cost > 110 or s in used_s or k in used_k:
                    continue
                pick[s] = imgs[k]
                used_s.add(s)
                used_k.add(k)
        # название модели: над таблицей строка с базовыми кодами («7414 7415»), над ней — имя серии («ZEO»)
        top = g[0]['word'][1]
        bw = [w for w in words if w[4].split()[0] == g[0]['base'] and w[3] < top - 2 and top - w[1] < 120]
        label = ''
        listed = []
        if bw:
            b0 = max(bw, key=lambda w: w[1])
            listed = [x for x in b0[4].split() if re.match(r'^\d{4}$', x)]
            nm = [w for w in words if abs(w[0] - b0[0]) < 60 and 3 < b0[1] - w[3] < 22 and not w[4].isdigit()]
            if nm:
                y = max(w[1] for w in nm)
                label = ' '.join(w[4] for w in sorted(nm, key=lambda w: w[0]) if abs(w[1] - y) < 3)
        bases = []
        for r in g:
            key = core.code_key(r['raw'])
            s = r['swatch']
            color = name_of.get(s[0], '') if s else ''
            if s and s[0] in pick:
                photo, verdict = pick[s[0]], 'фото цвета'
            elif main:
                one_color = len({x['swatch'][0] for x in g if x['swatch']}) <= 1
                photo, verdict = main, ('совпадает' if one_color else 'проверить цвет')
            else:
                continue
            out.append(dict(code=key, raw=r['raw'], label=(label + ' ' + color).strip()[:300], rect=photo,
                            colors={}, all_colors=[], shared=0, verdict=verdict))
            if r['base'] not in bases:
                bases.append(r['base'])
        for b in listed:
            if b not in bases:
                bases.append(b)
        if main:
            for b in bases:
                if any(it['code'] == b for it in out):
                    continue
                out.append(dict(code=b, raw=b, label=(label + ' (модель)').strip(), rect=main,
                                colors={}, all_colors=[], shared=0, verdict='фото модели'))
    return out
