#!/usr/bin/env python3
"""
Proforma / Invoice из PDF -> Excel-отчёт, с необязательной сверкой с прайс-листом.

Запуск:
    python invoice_reader.py ОТЧЁТ.xlsx ФАЙЛ1.pdf [ФАЙЛ2.pdf ...] [--price ПРАЙС1.xlsx --price ПРАЙС2.xlsx ...]

Что делает:
  1. В каждом PDF находит строку заголовков таблицы (Code / Description / Qty / Price / Amount,
     а также итальянские и русские варианты) и по её колонкам читает позиции.
  2. Проверяет каждую строку: кол-во × цена = сумма; сумму строк сравнивает с итогом документа.
  3. Если указан прайс — ищет каждый код в прайсе (точно, по короткому коду Bohemia или по началу
     кода, как у Guzzini: 1793.00 -> 179300156) и сравнивает цену.
"""
import os
import re
import sys
import bisect
import collections

import pymupdf
import xlsxwriter

import bohemia_photos as core

# ---------------------------------------------------------------- заголовки колонок
# Порядок важен: «Unit price» — это цена, «Total price» — сумма, «Item description» — описание.
ROLES = [
    ('ean', r'barcode|\bean\b|gtin|штрих'),
    ('vat', r'^(vat|iva|ндс|c\.?\s*iva|cod\.?\s*iva|tax)\b|(vat|iva|tax)\s*(rate|%|code|cod)'),
    ('total', r'amount|total|\bsum\b|value|importo|totale|сумма|стоимость|всего|^net$|^netto$'),
    ('price', r'price|prezzo|цена|\brate\b|unit\s*cost|\bcost\b|preis'),
    ('disc', r'disc|sconto|скидк|rabatt'),
    ('qty', r"q[\s'.\-]*t[yà]|quant|\bqta\b|\bpcs\b|кол|pieces|menge|\bpz\b|ordered|shipped|množ"),
    ('no', r'^(no\.?|#|№|n\.?|n°|pos\.?|поз\.?|line|№\s*п/п|п/п)$'),
    ('unit', r'^(unit|uom|u\.?\s?m\.?|u/m|ед\.?|ед\.\s*изм\.?|um|units?)$'),
    ('desc', r'desc|\bname\b|designation|descrizione|наимен|описан|goods|bezeichnung'),
    ('code', r'code|reference|materiale|\bart\b|art\.|artic|\bsku\b|\bref|codice|артик|код|\bitem\b|part\s*n|cat\.?\s*n|\bean\b'),
    ('desc', r'product|товар|prodotto|item'),
]
ROLE_RE = [(r, re.compile(p, re.I)) for r, p in ROLES]
STOP_RE = re.compile(r'^\s*(bank\b|iban\b|swift\b|aba\s+routing|celkem|summe|gesamt|sub\s*-?total|total|grand\s*total|итого|всего|totale|imponibile|amount\s+due|'
                     r'total\s+amount|net\s+total|to\s+pay|к\s+оплате)\b', re.I)
NOT_MONEY_RE = re.compile(r'weight|\bkg\b|brutto|netto|gross\s*w|net\s*w|\bвес|volume|\bm3\b|cbm|pallet|'
                          r'carton|boxes|колли|паллет|мест\b|pcs|шт', re.I)
NOISE_RE = re.compile(r'country\s+of\s+origin|custom\s*n|hs[\s-]*code|\bhst\b|made\s+in|delivery\s+date|'
                      r'confirmed\s+qty|pick\s+number|origine|\bpos:|\bean:|\bitf:|harmonized|mould\s+code|^schedule$|'
                      r'conf\s+qty|order\s+qty|n\.pall|\bkg[ln]:|pref\.\s*orig|nomenclatur|taric|\bnom\.', re.I)
NUM_IN_TEXT = re.compile(r'(?<![\w.,])\d{1,3}(?:[ \u00a0]\d{3})+(?:[.,]\d+)?(?![\w.,])|(?<![\w.,])\d[\d.,]*\d(?![\w])|(?<![\w.,])\d(?![\w.,])')
CODEISH_RE = re.compile(r'^(?=.*\d)[0-9A-Z][0-9A-Z./\-]{3,}$', re.I)


def role_of(text):
    t = text.strip().lower()
    if not t or len(t) > 40:
        return None
    for r, rx in ROLE_RE:
        if rx.search(t):
            return r
    return None


# ---------------------------------------------------------------- числа
def decimal_style(text):
    """',' если в документе цены вида 12,50; '.' если 12.50."""
    comma = len(re.findall(r'\d,\d{2}(?!\d)', text))
    dot = len(re.findall(r'\d\.\d{2}(?!\d)', text))
    return ',' if comma > dot else '.'


UNIT_TOK = re.compile(r'^(€|\$|£|₽|eur|usd|rub|pc|pcs|ea|nr|pz|zco|ct|kg|шт\.?|/|%|x)$', re.I)


def num_field(s, role, dec='.'):
    """Число из ячейки, где кроме числа есть единицы и валюта: «1 PC», «69.14 EUR /», «1 PC 69.14 EUR»."""
    v = parse_num(s, dec)
    if v is not None or not s:
        return v
    toks = str(s).split()
    nums = [parse_num(t, dec) for t in toks]
    if not all(n is not None or UNIT_TOK.match(t) for t, n in zip(toks, nums)):
        return None
    nums = [n for n in nums if n is not None]
    if not nums:
        return None
    return nums[0] if role == 'qty' else nums[-1]


def parse_num(s, dec='.'):
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = str(s).replace(' ', ' ').replace(' ', ' ').replace("'", '').strip()
    t = re.sub(r'(?i)€|\$|£|₽|eur|usd|rub|руб\.?|pcs\.?|шт\.?|pz\.?|nr\.?|%', '', t).strip()
    neg = t.startswith('-') or t.endswith('-') or (t.startswith('(') and t.endswith(')'))
    t = t.strip('()-+ ')
    if not t or not re.fullmatch(r'\d[\d .,]*', t):
        return None
    if ' ' in t and not re.fullmatch(r'\d{1,3}( \d{3})+([.,]\d+)?', t):
        return None                     # «3 45,00» — это два разных числа
    t = t.replace(' ', '')
    if ',' in t and '.' in t:
        if t.rfind(',') > t.rfind('.'):
            t = t.replace('.', '').replace(',', '.')
        else:
            t = t.replace(',', '')
    elif ',' in t:
        t = t.replace(',', '.') if (dec == ',' and t.count(',') == 1) else t.replace(',', '')
    elif t.count('.') == 1 and dec == ',' and re.fullmatch(r'\d{1,3}\.\d{3}', t):
        t = t.replace('.', '')          # 1.234 в документе с запятой — это тысячи
    elif t.count('.') > 1:
        t = t.replace('.', '')
    try:
        v = float(t)
    except ValueError:
        return None
    return -v if neg else v


# ---------------------------------------------------------------- строки страницы
def page_lines(page):
    """Слова страницы, сгруппированные в строки по вертикали."""
    words = [w for w in page.get_text('words') if w[4].strip()]
    words.sort(key=lambda w: ((w[1] + w[3]) / 2, w[0]))
    lines = []
    for w in words:
        yc = (w[1] + w[3]) / 2
        h = max(w[3] - w[1], 4)
        if lines and abs(lines[-1]['yc'] - yc) < h * 0.45:
            lines[-1]['words'].append(w)
        else:
            lines.append(dict(yc=yc, words=[w]))
    # строка чисел, сдвинутая на полстроки вниз (кол-во и сумма под ценой) — часть той же строки
    merged = []
    for ln in lines:
        if merged:
            a = merged[-1]
            ay0 = min(w[1] for w in a['words'])
            ah = max(w[3] for w in a['words']) - ay0
            by0 = min(w[1] for w in ln['words'])
            nums_only = all(NUMISH_RE.match(w[4]) or len(w[4]) <= 3 for w in ln['words'])
            no_overlap = all(w[2] <= v[0] or w[0] >= v[2] for w in ln['words'] for v in a['words'])
            if nums_only and no_overlap and 0 < by0 - ay0 < 0.7 * ah:
                a['words'] += ln['words']
                continue
        merged.append(ln)
    lines = merged
    for ln in lines:
        ln['words'].sort(key=lambda w: w[0])
        ln['y0'] = min(w[1] for w in ln['words'])
        ln['y1'] = max(w[3] for w in ln['words'])
        ln['text'] = ' '.join(w[4] for w in ln['words'])
    return lines


NUMISH_RE = re.compile(r'^[-+(]?[\d.,\u00a0]+[)%-]?$|^(€|\$|£|₽|eur|usd)$', re.I)


def make_cells(words, gap=7, split_nums=False):
    """Слова -> ячейки: соседние по горизонтали (или стоящие друг под другом) сливаются.
    split_nums: число не сливается с соседним словом-текстом («135,00 NI8C1» -> две ячейки)."""
    cells = []
    for w in sorted(words, key=lambda w: w[0]):
        if cells and w[0] <= cells[-1]['x1'] + gap:
            prev = cells[-1]['words'][-1][4]
            if split_nums and (bool(NUMISH_RE.match(prev)) != bool(NUMISH_RE.match(w[4]))) \
                    and w[0] > cells[-1]['x1'] + 3.5:
                cells.append(dict(x0=w[0], x1=w[2], words=[w]))
                continue
            c = cells[-1]
            c['x1'] = max(c['x1'], w[2])
            c['words'].append(w)
        else:
            cells.append(dict(x0=w[0], x1=w[2], words=[w]))
    for c in cells:
        c['text'] = ' '.join(w[4] for w in sorted(c['words'], key=lambda w: (round(w[1]), w[0])))
        c['role'] = role_of(c['text'])
    return cells


def is_header_line(ln):
    cells = make_cells(ln['words'], gap=4.5)
    roles = [c['role'] for c in cells if c['role']]
    return len(roles) >= 3 and len(roles) >= 0.6 * len(cells)


def find_header(lines):
    """Ищем строку заголовков (иногда в 2 строки: «Unit / price»). Возвращает (индекс последней строки, колонки)."""
    best = None
    for i, ln in enumerate(lines):
        for span in (1, 2, 3):
            if i + span > len(lines):
                break
            group = lines[i:i + span]
            if span > 1 and group[-1]['y0'] - group[-2]['y1'] > 8:
                break
            cells = make_cells([w for g in group for w in g['words']], gap=4.5)
            roles = [c['role'] for c in cells if c['role']]
            if len(cells) < 3 or len(roles) < 3 or len(roles) < 0.6 * len(cells):
                continue
            rs = set(roles)
            if not (rs & {'code', 'desc'}) or not (rs & {'qty', 'price', 'total'}):
                continue
            score = len(rs)
            if not best or score > best[0]:
                best = (score, i + span - 1, cells)
        if best and i > best[1] + 2:
            break
    if not best:
        return None
    _, idx, cells = best
    # одна роль — одна колонка: повтор (например, две «Price») считаем дополнительной колонкой
    seen = collections.Counter()
    cols = []
    for c in cells:
        r = c['role'] or 'other'
        seen[r] += 1
        cols.append(dict(role=r if seen[r] == 1 else f'{r}{seen[r]}', x0=c['x0'], x1=c['x1'], title=c['text']))
    return idx, cols


SINGLE_ROLES = {'no', 'code', 'qty', 'unit', 'price', 'disc', 'total', 'vat', 'ean'}


def assign(words, cols):
    """Раскладываем строку по колонкам заголовка. Сначала слова сливаем в ячейки по промежуткам
    (длинное описание остаётся одной ячейкой), затем ячейку отдаём колонке с наибольшим перекрытием.
    Если в «однозначную» колонку (кол-во, цена, сумма…) попало две ячейки — лишнюю сдвигаем в соседнюю
    пустую колонку: числа часто выровнены вправо и стоят не под своим заголовком."""
    placed = collections.defaultdict(list)   # индекс колонки -> ячейки
    cells = []
    for c in make_cells(words, gap=5, split_nums=True):
        # ячейка накрывает заголовки двух колонок («100105-004 Bamboo…» под ITEM NO. и DESCRIPTION) — делим по словам
        hit = [col for col in cols if min(c['x1'], col['x1']) - max(c['x0'], col['x0']) > 0]
        if len(hit) > 1 and len(c['words']) > 1:
            groups = []
            for w in c['words']:
                k = max(range(len(cols)), key=lambda j: (min(w[2], cols[j]['x1']) - max(w[0], cols[j]['x0']),
                                                         -abs((w[0] + w[2]) / 2 - (cols[j]['x0'] + cols[j]['x1']) / 2)))
                if min(w[2], cols[k]['x1']) - max(w[0], cols[k]['x0']) <= 0 and groups:
                    k = groups[-1][0]
                if groups and groups[-1][0] == k:
                    groups[-1][1].append(w)
                else:
                    groups.append((k, [w]))
            for _, ws in groups:
                cells.append(dict(x0=ws[0][0], x1=max(w[2] for w in ws), words=ws, text=' '.join(w[4] for w in ws)))
        else:
            cells.append(c)
    for c in cells:
        best, ov, dist = None, 0, None
        for j, col in enumerate(cols):
            o = min(c['x1'], col['x1']) - max(c['x0'], col['x0'])
            d = max(col['x0'] - c['x1'], c['x0'] - col['x1'], 0)
            if o > ov or (ov <= 0 and o <= 0 and (dist is None or d < dist)):
                best, ov, dist = j, max(o, ov), d
        placed[best].append(c)
    for _ in range(3):
        moved = False
        for j in sorted(placed):
            cells = placed[j]
            if len(cells) < 2 or cols[j]['role'].rstrip('0123456789') not in SINGLE_ROLES:
                continue
            col = cols[j]
            cxc = (col['x0'] + col['x1']) / 2
            keep = min(cells, key=lambda c: max(col['x0'] - c['x1'], c['x0'] - col['x1'], 0) * 1000
                       + abs((c['x0'] + c['x1']) / 2 - cxc))
            for c in [c for c in cells if c is not keep]:
                step = -1 if c['x1'] <= keep['x0'] else 1
                k = j + step
                if 0 <= k < len(cols) and not placed.get(k):
                    placed[j].remove(c)
                    placed[k] = [c]
                    moved = True
        if not moved:
            break
    out = collections.defaultdict(list)
    for j in sorted(placed):
        for c in sorted(placed[j], key=lambda c: c['x0']):
            out[cols[j]['role']].append(c['text'])
    return {r: ' '.join(v) for r, v in out.items()}


def h_rules(page):
    """Y горизонтальных линий таблицы (границы строк), если они нарисованы."""
    ys = []
    try:
        for d in page.get_drawings():
            for it in d['items']:
                if it[0] == 'l':
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 1 and abs(p1.x - p2.x) > 40:
                        ys.append(p1.y)
                elif it[0] == 're':
                    r = it[1]
                    if r.height < 2 and r.width > 40:
                        ys.append((r.y0 + r.y1) / 2)
                    elif r.width > 40 and r.height > 2:
                        ys += [r.y0, r.y1]
    except Exception:
        pass
    return sorted(ys)


def rule_between(rules, y_top, y_bottom):
    j = bisect.bisect_right(rules, y_top - 0.5)
    return j < len(rules) and rules[j] < y_bottom + 0.5


# ---------------------------------------------------------------- документ
def doc_info(text):
    low = text.lower()
    if re.search(r'pro\s*-?\s*forma|pre\s*-?\s*fattura', low):
        kind = 'Proforma'
    elif re.search(r'order\s+confirmation|conferma\s+d.?ordine|sales\s+order|\border\b|ordine|заказ', low):
        kind = 'Order'
    elif re.search(r'invoice|fattura|счет|счёт|инвойс|rechnung', low):
        kind = 'Invoice'
    else:
        kind = ''
    num = ''
    m = re.search(r'(?:pro\s*-?\s*forma(?:\s+(?:invoice|faktura))?|invoice|fattura|order\s+confirmation|'
                  r'order|ordine|счет|счёт|инвойс)\s*'
                  r'(?:no\.?|nr\.?|n\.|n°|№|#|number|num\.?)?\s*[:.]?\s*(?:[A-Z]{1,3}\d?[A-Z]?\s+)?'
                  r'((?=[A-Z0-9/\-.]*\d{3})[A-Z0-9][A-Z0-9/\-.]*)',
                  text, re.I)
    if m:
        num = m.group(1).strip('.-')
    m2 = re.search(r'([A-Z0-9][A-Z0-9/\-.]*\d[A-Z0-9/\-.]*)\s*\n\s*N\.?\s*DOC', text)   # «OM26-06703 / N.DOC.»
    if m2:
        num = m2.group(1)
    d = re.search(r'\b(\d{1,2}[./-]\d{1,2}[./-](?:\d{4}|\d{2})|\d{4}-\d{2}-\d{2})\b', text)
    cur = collections.Counter()
    cur['EUR'] = len(re.findall(r'€|\bEUR\b|\beuro\b', text, re.I))
    cur['USD'] = len(re.findall(r'\$|\bUSD\b', text))
    cur['RUB'] = len(re.findall(r'₽|\bRUB\b|руб', text, re.I))
    currency = cur.most_common(1)[0][0] if max(cur.values()) else ''
    return dict(kind=kind, number=num, date=d.group(1) if d else '', currency=currency)


class InvoiceError(RuntimeError):
    """Понятная пользователю ошибка разбора документа."""


def read_invoice(path, log=print):
    """Возвращает (сведения о документе, позиции)."""
    doc = pymupdf.open(path)
    full = '\n'.join(p.get_text() for p in doc)
    if len(full.strip()) < 30:
        raise InvoiceError(f'{os.path.basename(path)}: в PDF нет текста (похоже на скан). '
                           f'Нужен PDF, выгруженный из программы, а не отсканированный.')
    dec = decimal_style(full)
    info = doc_info(full)
    info['file'] = os.path.basename(path)
    items, totals = [], []
    cols = None
    for pn, page in enumerate(doc, 1):
        lines = page_lines(page)
        hdr = find_header(lines)
        if hdr:
            start, cols = hdr[0] + 1, hdr[1]
            strict = False
        elif cols:
            start, strict = 0, True   # продолжение таблицы без заголовка на новой странице
        else:
            continue
        stopped = False
        last, pending = None, []      # pending: строки только с текстом (перенос описания)
        rules = h_rules(page)

        def flush(to_next=None):
            """Отдаём накопленные строки описания предыдущей или следующей позиции."""
            if not pending:
                return
            target = last
            if to_next is not None:
                if last is None:
                    target = None     # текст до первой позиции — примечания, а не описание
                elif rules and rule_between(rules, last['_y1'], pending[0]['y0']) \
                        and not rule_between(rules, pending[-1]['y1'], to_next['_y0']):
                    target = to_next
                elif not (rules and rule_between(rules, pending[-1]['y1'], to_next['_y0'])):
                    gap_prev = pending[0]['y0'] - last['_y1']
                    gap_next = to_next['_y0'] - pending[-1]['y1']
                    if gap_next < gap_prev - 1:
                        target = to_next
            if target is not None:
                extra = ' '.join(p['text'] for p in pending)
                if target is to_next:
                    target['desc'] = (extra + ' ' + target['desc']).strip()
                else:
                    target['desc'] = (target['desc'] + ' ' + extra).strip()
            pending.clear()

        for ln in lines[start:]:
            if STOP_RE.search(ln['text']) or (stopped and re.search(r'\btotal|итого|всего', ln['text'], re.I)):
                if not stopped:
                    flush()
                stopped = True
                if not NOT_MONEY_RE.search(ln['text']):
                    nums = [parse_num(m, dec) for m in NUM_IN_TEXT.findall(ln['text'])]
                    nums = [x for x in nums if x is not None]
                    if nums:
                        totals.append((ln['text'], nums[-1]))
                continue
            if stopped:
                continue
            if is_header_line(ln):
                continue          # вторая строка шапки (например, английская под итальянской)
            f = assign(ln['words'], cols)
            if 'price' in f and parse_num(f['price'], dec) is None:
                m = re.search(r'(\d+)\s*[A-Za-z]{0,4}\s*[xX×]\s*([\d.,]+)', f['price'])
                if m:   # «100 ZCO X 54,67 EUR» — цена за 100 штук
                    per, val = parse_num(m.group(1), dec), parse_num(m.group(2), dec)
                    if per and val is not None:
                        f['price'] = str(round(val / per, 6)).replace('.', dec)
            spill = []
            for r in list(f):
                if r.rstrip('0123456789') in ('qty', 'price', 'total') and \
                        num_field(f[r], r.rstrip('0123456789'), dec) is None:
                    spill.append(f.pop(r))        # текст попал в числовую колонку — это описание
            if spill:
                f['other_spill'] = ' '.join(spill)
            qty, price, total = (next((v for v in (num_field(f.get(r + s), r, dec) for s in ('', '2', '3'))
                                       if v is not None), None) for r in ('qty', 'price', 'total'))
            code = (f.get('code') or '').strip()
            others = ' '.join(v for r, v in f.items() if r.startswith('other') and not NUMISH_RE.match(v))
            desc = ' '.join(x for x in [(f.get(r) or '').strip() for r in ('desc', 'desc2', 'desc3')] + [others] if x)
            no = f.get('no', '')
            if not code and desc:
                toks = desc.split()
                if len(toks) > 1 and re.fullmatch(r'\d{1,3}', toks[0]) and CODEISH_RE.match(toks[1]):
                    no, toks = no or toks[0], toks[1:]      # «1 91L/1SC69/…» — номер строки и код
                if 'code' not in {c['role'] for c in cols} and CODEISH_RE.match(toks[0]):
                    code, toks = toks[0], toks[1:]
                desc = ' '.join(toks)
            has_num = qty is not None or price is not None or total is not None
            if not code and re.match(r'^[^\W\d][\w .\-/]{1,24}:(\s|$)', desc):
                continue          # «Price: 76.82 EUR / 1 PC» — служебная строка под позицией
            enough = (qty is not None and (price is not None or total is not None)) or \
                (price is not None and total is not None)
            if has_num and (code or desc) and enough and (not strict or qty is not None):
                it = dict(file=info['file'], page=pn, no=no, code=code, desc=desc,
                          qty=qty, unit=f.get('unit', ''), price=price, disc=f.get('disc', ''), total=total,
                          _y0=ln['y0'], _y1=ln['y1'])
                flush(it)
                last = it
                items.append(it)
            elif NOISE_RE.search(ln['text']):
                continue
            elif not has_num and (desc or code) and (last is not None or not strict):
                if pending and ln['y0'] - pending[-1]['y1'] > 25:
                    pending.clear()
                if last is not None and not pending and ln['y0'] - last['_y1'] > 40:
                    continue
                pending.append(dict(text=' '.join(x for x in (code, desc) if x), y0=ln['y0'], y1=ln['y1']))
            elif has_num:
                flush()
        flush()
        if stopped:
            cols = None if not hdr else cols
    # колонка «Item» оказалась номером строки (10, 20, 30…), а код — первым словом описания
    if items and all(re.fullmatch(r'\d{1,3}', it['code'] or '') for it in items):
        firsts = [it['desc'].split()[0] if it['desc'] else '' for it in items]
        if sum(bool(CODEISH_RE.match(x)) for x in firsts) >= 0.8 * len(items):
            for it, first in zip(items, firsts):
                it['no'] = it['no'] or it['code']
                if CODEISH_RE.match(first):
                    it['code'], it['desc'] = first, it['desc'][len(first):].strip()
                else:
                    it['code'] = ''
    for it in items:
        it.pop('_y0', None)
        it.pop('_y1', None)
        it['check'] = line_check(it, dec)
    info['items_sum'] = round(sum(it['total'] or 0 for it in items), 2)
    info['totals'] = totals
    info['doc_total'] = pick_total(totals, info['items_sum'])
    log(f'  {info["file"]}: {info["kind"] or "документ"} {info["number"]} — позиций: {len(items)}')
    if not items:
        log('    ! не нашёл таблицу позиций (нет строки заголовков Code / Description / Qty / Price)')
    return info, items


def read_table_file(path, log=print):
    """Документ в Excel/CSV (выгрузка заказа): ищем строку заголовков и читаем позиции."""
    rows, _ = core.read_price(path)
    name = os.path.basename(path)
    info = dict(file=name, kind='', number='', date='', currency='', totals=[])
    head = None
    for i, row in enumerate(rows[:40]):
        roles = {}
        for c, v in enumerate(row):
            r = role_of(str(v)) if v not in (None, '') else None
            if r and r not in roles:
                roles[r] = c
        if len(roles) >= 3 and (set(roles) & {'code', 'desc'}) and (set(roles) & {'qty', 'price', 'total'}):
            head = (i, roles)
            break
    items = []
    if head:
        i, roles = head
        if 'code' not in roles:   # «Materiale», «Reference» и т.п.
            for c, v in enumerate(rows[i]):
                if v and re.search(r'materiale|material|reference|riferimento', str(v), re.I):
                    roles['code'] = c
                    break
        for rn, row in enumerate(rows[i + 1:], i + 2):
            def g(r):
                c = roles.get(r)
                return row[c] if c is not None and c < len(row) else None
            code = g('code')
            if isinstance(code, float) and code.is_integer():
                code = int(code)
            code = str(code).strip() if code not in (None, '') else ''
            desc = str(g('desc') or '').strip()
            qty, price, total = (parse_num(g(r)) for r in ('qty', 'price', 'total'))
            if not (code or desc) or (qty is None and price is None and total is None):
                continue
            if total is None and qty is not None and price is not None:
                total = round(qty * price, 2)
            items.append(dict(file=name, page='', no=str(g('no') or ''), code=code, desc=desc, qty=qty,
                              unit=str(g('unit') or ''), price=price, disc=str(g('disc') or ''),
                              total=round(total, 2) if total is not None else None))
    for it in items:
        it['check'] = line_check(it, '.')
    info['items_sum'] = round(sum(it['total'] or 0 for it in items), 2)
    info['doc_total'] = None
    log(f'  {name}: позиций: {len(items)}')
    if not items:
        log('    ! не нашёл таблицу позиций (нет строки заголовков Code / Description / Qty / Price)')
    return info, items


def line_check(it, dec):
    q, p, t = it['qty'], it['price'], it['total']
    if q is None or p is None or t is None:
        return ''
    base = q * p
    d = parse_num(it.get('disc'), dec) if it.get('disc') else None
    variants = [base]
    if d:
        variants += [base * (1 - d / 100), base - d, q * (p - d)]
    tol = max(0.02, 0.005 * abs(t))
    return 'ок' if any(abs(v - t) <= tol for v in variants) else 'проверить'


def pick_total(totals, items_sum):
    """Итог документа: сначала тот, что совпал с суммой строк, иначе последний «Total»."""
    for _, v in totals:
        if abs(v - items_sum) <= 0.05:
            return v
    for text, v in reversed(totals):
        if re.search(r'grand|к оплате|to pay|due|totale|итого', text, re.I):
            return v
    return totals[-1][1] if totals else None


# ---------------------------------------------------------------- прайс
CODE_HEAD = re.compile(r'code|codice|артик|код|\bsap\b|\bart\b|art\.|\bitem\b|\bsku\b|\bref|production\s*n', re.I)
PRICE_HEAD = re.compile(r'^(eur|usd|rub|€|\$|₽)\b|price|prezzo|цена|listino|preis', re.I)
DESC_HEAD = re.compile(r'desc|наимен|описан|\bname\b|designation|descrizione', re.I)


def header_currency(title):
    t = title.upper()
    for cur, rx in (('EUR', r'EUR|€'), ('USD', r'USD|\$'), ('RUB', r'RUB|₽|РУБ')):
        if re.search(rx, t):
            return cur
    return ''


def read_pdf_price(path, log=print):
    """PDF-прайс (таблица «код / название / … / цена» без количества) -> строки как у листа Excel:
    первая строка — заголовки, дальше по строке на товар. Заголовок ищется на каждой странице."""
    import pymupdf
    doc = pymupdf.open(path)
    head = ['Item No.', 'Description', 'Size', 'PU', 'Price EUR']
    out = [tuple(head)]
    cur = None
    for page in doc:
        lines = page_lines(page)
        h = find_header(lines)
        start = 0
        if h:
            start, cols = h[0] + 1, h[1]
            cur = cols
        if not cur:
            continue
        roles = [c['role'] for c in cur]
        if 'code' not in roles or not ({'price', 'total'} & set(roles)):
            continue
        pcol = 'price' if 'price' in roles else 'total'
        extra = [r for r in roles if r not in ('code', 'desc', pcol)]
        for ln in lines[start:]:
            if STOP_RE.search(ln['text']) and not re.search(r'\d', ln['words'][0][4]):
                continue
            row = assign(ln['words'], cur)
            code = (row.get('code') or '').strip()
            pr = row.get(pcol) or ''
            m = re.search(r'-?\d[\d.\s\u00a0]*,\d{1,4}|-?\d[\d,\s\u00a0]*\.\d{1,4}|^\d+$', pr.replace('€', '').strip())
            if not code or not m or not re.search(r'\d', code) or len(code) > 30:
                continue
            val = parse_num(m.group(0).strip(), ',' if ',' in m.group(0) and '.' not in m.group(0) else '.')
            if val is None:
                continue
            ex = [row.get(r, '') for r in extra][:2]
            ex += [''] * (2 - len(ex))
            out.append((code, (row.get('desc') or '').strip(), ex[0], ex[1], val))
    log(f'  PDF-прайс «{os.path.basename(path)}»: строк с ценой {len(out) - 1}')
    return out


class PriceList:
    """Прайс: индекс кодов (точный, короткий Bohemia, по началу кода для Guzzini) и цены."""

    def __init__(self, paths, log=print):
        """paths — один прайс или список прайсов (разные бренды проверяются вместе)."""
        if isinstance(paths, str):
            paths = [paths]
        self.rows = []
        self.index = collections.defaultdict(list)
        self.price_cols = []      # (заголовок, валюта)
        for path in paths:
            self.name = os.path.basename(path)
            before = len(self.rows)
            for title, rows in self._sheets(path):
                self._add_sheet(title, rows, log)
            if len(self.rows) == before:
                raise core.PriceError(f'В прайсе «{self.name}» не нашёл заголовков: нужна колонка кода '
                                      f'(CODE / Item / Артикул / Код) и колонка цены (EUR / USD / Price / Цена).')
            log(f'  прайс «{self.name}»: строк с кодами {len(self.rows) - before}')
        self.keys = sorted(self.index)

    @staticmethod
    def price_of(rec, currency):
        """(колонка, цена) строки прайса: колонка в валюте документа, иначе первая."""
        if not rec['prices']:
            return '', None
        for title, v in rec['prices'].items():
            if currency and header_currency(title) == currency:
                return title, v
        title = next(iter(rec['prices']))
        return title, rec['prices'][title]

    @staticmethod
    def _sheets(path):
        """Все листы .xlsx (в прайсах Guzzini товары разложены по листам), иначе первый лист."""
        if path.lower().endswith('.pdf'):
            return [('', read_pdf_price(path, log=lambda *a: None))]
        if path.lower().endswith(('.xlsx', '.xlsm')):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                return [(ws.title, [tuple(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
            except Exception:
                pass
        rows, _ = core.read_price(path)
        return [('', rows)]

    def _add_sheet(self, sheet, rows, log):
        head = None
        for i, row in enumerate(rows[:30]):
            texts = [str(v).strip() if v is not None else '' for v in row]
            code_cols = [c for c, t in enumerate(texts) if t and CODE_HEAD.search(t) and not PRICE_HEAD.search(t)]
            price_cols = [c for c, t in enumerate(texts) if t and PRICE_HEAD.search(t) and len(t) < 40]
            if code_cols and price_cols:
                desc = next((c for c, t in enumerate(texts) if t and DESC_HEAD.search(t)), None)
                head = (i, code_cols, price_cols, desc, texts)
                break
        if not head:
            return
        i, code_cols, price_cols, desc_col, titles = head
        # для показа берём «главную» колонку кода (CODE / Item / Codice / Артикул), а не SAP, EAN и т.п.
        code_cols.sort(key=lambda c: 0 if re.match(r'(code|item|codice|код|артикул|art)', titles[c].strip(), re.I)
                       and not re.search(r'ean|upc|bar', titles[c], re.I) else 1)
        for c in price_cols:
            pc = (titles[c], header_currency(titles[c]))
            if pc not in self.price_cols:
                self.price_cols.append(pc)
        for rn, row in enumerate(rows[i + 1:], i + 2):
            codes = []
            for c in code_cols:
                v = row[c] if c < len(row) else None
                if v in (None, ''):
                    continue
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                s = str(v).strip()
                # «STO 0141», «EGL 7004» (запчасти Helios) — код целиком; иначе первое слово
                if not re.match(r'^[A-Za-z]{1,4}\s+\d[\w./-]*$', s):
                    s = s.split()[0] if s else ''
                if s and not s.startswith('='):
                    codes.append(s)
            if not codes:
                continue
            prices = {}
            for c in price_cols:
                v = row[c] if c < len(row) else None
                pv = parse_num(v, ',' if isinstance(v, str) and re.search(r'\d,\d{2}$', v) else '.')
                if pv is not None:
                    prices[titles[c]] = round(pv, 4)
            rec = dict(row=rn, sheet=sheet, file=self.name, code=codes[0], codes=codes, prices=prices,
                       desc=str(row[desc_col] or '').strip() if desc_col is not None and desc_col < len(row) else '')
            self.rows.append(rec)
            keys = set()
            for s in codes:
                k = core.code_key(core.clean_code(s))
                if k:
                    keys.add(k)
                sc = core.short_code(core.clean_code(s))
                if sc:
                    keys.add(core.code_key(sc))
            for k in keys:
                self.index[k].append(rec)

    def price_col(self, currency):
        for title, cur in self.price_cols:
            if currency and cur == currency:
                return title
        return self.price_cols[0][0]

    def lookup(self, code):
        """-> (список строк прайса, как найдено)."""
        k = core.code_key(code)
        if len(k) < 3:
            return [], ''
        if k in self.index:
            rows = self.index[k]
            exact = any(core.code_key(core.clean_code(c)) == k for r in rows for c in r['codes'])
            return rows, 'точно' if exact else 'по короткому коду'
        sc = core.short_code(core.clean_code(code))
        if sc and core.code_key(sc) in self.index:
            return self.index[core.code_key(sc)], 'по короткому коду'
        # 2944-C-001 (код с цветом) -> 2944-C -> 2944 в прайсе (Helios)
        parts = str(code).strip().split('-')
        if len(parts) > 1 and re.match(r'^[A-Za-z]{0,4}\s?\d{3,}$', parts[0]) and all(re.match(r'^\w{1,4}$', p) for p in parts[1:]):
            for n in range(len(parts) - 1, 0, -1):
                kb = core.code_key('-'.join(parts[:n]))
                if kb in self.index:
                    return self.index[kb], 'по коду без цвета'
        if len(k) >= 5:
            # код документа — начало кода прайса (1793.00 -> 179300156)
            j = bisect.bisect_left(self.keys, k)
            found = []
            # остаток — код цвета/упаковки (до 5 знаков); так коды других брендов не цепляются случайно
            while j < len(self.keys) and self.keys[j].startswith(k) and len(self.keys[j]) - len(k) <= 5:
                found += [r for r in self.index[self.keys[j]] if r not in found]
                j += 1
            if found:
                return found, 'по началу кода'
            # код прайса — начало кода документа
            for n in range(len(k) - 1, max(4, len(k) - 6), -1):
                if k[:n] in self.index:
                    return self.index[k[:n]], 'по началу кода'
        return [], ''


# ---------------------------------------------------------------- отчёт
FILLS = {'ок': '#C6EFCE', 'проверить': '#FFEB9C', 'цена совпадает': '#C6EFCE', 'цена выше прайса': '#FFC7CE',
         'цена ниже прайса': '#FFEB9C', 'нет цены в прайсе': '#FFEB9C', 'нет в прайсе': '#FFC7CE',
         'несколько в прайсе': '#FFEB9C', 'сходится': '#C6EFCE', 'не сходится': '#FFC7CE'}


def compare(items, infos, price, log=print):
    stats = collections.Counter()
    cur_of = {i['file']: i['currency'] for i in infos}
    for it in items:
        rows, how = price.lookup(it['code']) if it['code'] else ([], '')
        if not rows and it['desc']:
            # код иногда стоит в начале описания
            first = it['desc'].split()[0]
            if CODEISH_RE.match(first):
                rows, how = price.lookup(first)
        cur = cur_of.get(it['file'], '')
        if not rows:
            it.update(pl_file='', price_col='', pl_code='', pl_desc='', pl_price=None, diff=None, pct=None,
                      how='', status='нет в прайсе')
        else:
            r = rows[0]
            col, plp = price.price_of(r, cur)
            shown = list(dict.fromkeys(x['code'] for x in rows))
            it.update(pl_file=r['file'], price_col=col,
                      pl_code=', '.join(shown[:5]) + (' …' if len(shown) > 5 else ''),
                      pl_desc=r['desc'], pl_price=plp, how=how)
            distinct = {price.price_of(x, cur)[1] for x in rows}
            if plp is None:
                it.update(diff=None, pct=None, status='нет цены в прайсе')
            elif it['price'] is None:
                it.update(diff=None, pct=None, status='проверить')
            else:
                diff = round(it['price'] - plp, 4)
                pct = round(100 * diff / plp, 2) if plp else None
                tol = max(0.01, 0.005 * plp)
                st = 'цена совпадает' if abs(diff) <= tol else ('цена выше прайса' if diff > 0 else 'цена ниже прайса')
                it.update(diff=diff, pct=pct, status=st)
            if len(rows) > 1 and len(distinct) > 1:
                it['status'] = 'несколько в прайсе'
        stats[it['status']] += 1
    for k, v in stats.most_common():
        log(f'  {k}: {v}')
    return stats


def write_report(path, infos, items, price=None, order_rows=None, order_items=None):
    wb = xlsxwriter.Workbook(path)
    base = dict(valign='top', text_wrap=True)
    hdr = wb.add_format(dict(base, bold=True, bg_color='#F2F2F2', border=1))
    cell = wb.add_format(base)
    money = wb.add_format(dict(base, num_format='#,##0.00'))
    colored = {}

    def fill(status, color=None):
        color = color or FILLS.get(status)
        if not color:
            return cell
        return colored.get(color) or colored.setdefault(color, wb.add_format(dict(base, bg_color=color)))

    def put(ws, r, vals, fmts):
        for c, (v, f) in enumerate(zip(vals, fmts)):
            if v is None or v == '':
                ws.write_blank(r, c, None, f)
            else:
                ws.write(r, c, v, f)

    if order_rows is not None:
        import orders as om
        wo = wb.add_worksheet('Сверка с заказом')
        oc = [('Код (заказ)', 20), ('Описание', 38), ('Кол-во в заказе', 10), ('Цена в заказе', 11),
              ('Сумма в заказе', 12), ('Код в документе', 20), ('Кол-во в документе', 11), ('Цена в документе', 11),
              ('Сумма в документе', 12), ('Разница кол-во', 10), ('Разница цены', 10), ('Расхождение', 30),
              ('Файл заказа', 26), ('Документ поставщика', 26)]
        for c, (t, w) in enumerate(oc):
            wo.set_column(c, c, w)
        wo.write_row(0, 0, [t for t, _ in oc], hdr)
        for r, x in enumerate(order_rows, 1):
            dq = None if x['o_qty'] is None or x['d_qty'] is None else round(x['d_qty'] - x['o_qty'], 4)
            dp = None if x['o_price'] is None or x['d_price'] is None else round(x['d_price'] - x['o_price'], 4)
            if x['status'].startswith('совпадает ('):    # в документе упаковками — разница только в единицах
                dq = dp = None
            put(wo, r, [x['code'], x['desc'], x['o_qty'], x['o_price'], x['o_sum'], x['d_code'], x['d_qty'],
                        x['d_price'], x['d_sum'], dq, dp, x['status'], x['o_files'], x['d_files']],
                [cell, cell, cell, money, money, cell, cell, money, money, cell, money,
                 fill(x['status'], om.fill_of(x['status'])), cell, cell])
        n = len(order_rows) + 2
        o_sum = round(sum(x['o_sum'] or 0 for x in order_rows), 2)
        d_sum = round(sum(x['d_sum'] or 0 for x in order_rows), 2)
        wo.write(n, 1, 'Итого', hdr)
        wo.write(n, 4, o_sum, money)
        wo.write(n, 8, d_sum, money)
        wo.freeze_panes(1, 0)
        wo.autofilter(0, 0, max(len(order_rows), 1), len(oc) - 1)

    if order_items:
        wz = wb.add_worksheet('Заказ')
        zc = [('Файл заказа', 26), ('Лист', 10), ('Строка', 7), ('Код', 20), ('Описание', 40), ('Кол-во', 9),
              ('Цена', 11), ('Сумма', 12)]
        if price:
            zc += [('Прайс', 22), ('Код в прайсе', 22), ('Цена прайса', 11), ('Разница', 10), ('Разница %', 9),
                   ('Сверка с прайсом', 18)]
        for c, (t, w) in enumerate(zc):
            wz.set_column(c, c, w)
        wz.write_row(0, 0, [t for t, _ in zc], hdr)
        for r, it in enumerate(order_items, 1):
            vals = [it['file'], it['sheet'], it['row'], it['code'], it['desc'], it['qty'], it['price'], it['total']]
            fmts = [cell] * 6 + [money, money]
            if price:
                vals += [it.get('pl_file'), it.get('pl_code'), it.get('pl_price'), it.get('diff'), it.get('pct'),
                         it.get('status')]
                fmts += [cell, cell, money, money, cell, fill(it.get('status'))]
            put(wz, r, vals, fmts)
        wz.freeze_panes(1, 0)

    if not items and order_rows is None and order_items:
        wb.close()
        return
    ws = wb.add_worksheet('Позиции')
    cols = [('Файл', 26), ('Стр.', 5), ('№', 5), ('Код', 18), ('Описание', 40), ('Кол-во', 9), ('Ед.', 6),
            ('Цена', 11), ('Скидка', 8), ('Сумма', 12), ('Кол-во×цена', 12)]
    if price:
        cols += [('Прайс', 22), ('Код в прайсе', 22), ('Как найден', 14), ('Описание (прайс)', 34), ('Колонка цены', 10),
                 ('Цена прайса', 11), ('Разница', 10), ('Разница %', 9), ('Сверка с прайсом', 18)]
    for c, (t, w) in enumerate(cols):
        ws.set_column(c, c, w)
    ws.write_row(0, 0, [t for t, _ in cols], hdr)
    for r, it in enumerate(items, 1):
        vals = [it['file'], it['page'], it['no'], it['code'], it['desc'], it['qty'], it['unit'],
                it['price'], it['disc'], it['total'], it['check']]
        fmts = [cell, cell, cell, cell, cell, cell, cell, money, cell, money, fill(it['check'])]
        if price:
            vals += [it['pl_file'], it['pl_code'], it['how'], it['pl_desc'], it['price_col'], it['pl_price'], it['diff'],
                     it['pct'], it['status']]
            fmts += [cell, cell, cell, cell, cell, money, money, cell, fill(it['status'])]
        for c, (v, f) in enumerate(zip(vals, fmts)):
            if v is None or v == '':
                ws.write_blank(r, c, None, f)
            else:
                ws.write(r, c, v, f)
    ws.freeze_panes(1, 0)
    ws.autofilter(0, 0, max(len(items), 1), len(cols) - 1)

    ws2 = wb.add_worksheet('Документы')
    cols2 = [('Файл', 30), ('Тип', 10), ('Номер', 16), ('Дата', 11), ('Валюта', 7), ('Позиций', 8),
             ('Сумма строк', 13), ('Итог в документе', 14), ('Сверка итога', 13), ('Строки «Total» в документе', 50)]
    for c, (t, w) in enumerate(cols2):
        ws2.set_column(c, c, w)
    ws2.write_row(0, 0, [t for t, _ in cols2], hdr)
    for r, i in enumerate(infos, 1):
        n = sum(1 for it in items if it['file'] == i['file'])
        dt = i['doc_total']
        chk = '' if dt is None else ('сходится' if abs(dt - i['items_sum']) <= 0.05 else 'не сходится')
        vals = [i['file'], i['kind'], i['number'], i['date'], i['currency'], n, i['items_sum'], dt, chk,
                '\n'.join(f'{t}' for t, _ in i['totals'][:8])]
        for c, v in enumerate(vals):
            f = money if c in (6, 7) else (fill(chk) if c == 8 else cell)
            if v is None or v == '':
                ws2.write_blank(r, c, None, f)
            else:
                ws2.write(r, c, v, f)
    ws2.freeze_panes(1, 0)
    wb.close()


def run(pdfs, out_xlsx, price_path=None, log=print, progress=None, order_paths=None):
    """pdfs — список PDF/Excel (Proforma/Invoice), out_xlsx — путь отчёта,
    price_path — прайс или список прайсов (необязательно),
    order_paths — наши заказы поставщику (Excel, необязательно): сверка «заказ ↔ документы поставщика»."""
    pdfs = list(pdfs or [])
    if not pdfs and not order_paths:
        raise InvoiceError('Не выбраны ни документы поставщика (Proforma / Invoice), ни заказы.')
    infos, items, errors = [], [], []
    log('Читаю документы...')
    for k, p in enumerate(pdfs):
        if progress:
            progress(0.7 * k / len(pdfs))
        try:
            if p.lower().endswith(('.xlsx', '.xls', '.xlsm', '.csv')):
                info, its = read_table_file(p, log)
            else:
                info, its = read_invoice(p, log)
        except (InvoiceError, core.PriceError) as e:
            log(f'  ! {e}')
            errors.append(str(e))
            continue
        infos.append(info)
        items += its
    if pdfs and not items:
        raise InvoiceError('Ни в одном документе не нашёл позиций. ' + (errors[0] if errors else
                           'Нужна таблица со строкой заголовков вроде Code / Description / Qty / Price / Amount.'))
    order_items, order_rows, ostats = [], None, collections.Counter()
    if order_paths:
        import orders as om
        log('Читаю заказы...')
        for p in order_paths:
            try:
                order_items += om.read_order(p, log)
            except Exception as e:
                log(f'  ! {os.path.basename(p)}: {e}')
                errors.append(str(e))
        if not order_items:
            raise InvoiceError('В файлах заказов не нашёл позиций с количеством '
                               '(нужны колонки Reference / Code и Quantity / Кол-во).')
        if items:
            log('Сверяю заказ с документами поставщика...')
            order_rows = om.compare_orders(order_items, items)
            for x in order_rows:
                ostats[x['status']] += 1
            for k, v in ostats.most_common():
                log(f'  {k}: {v}')
    stats = collections.Counter()
    price = None
    if price_path:
        log('Читаю прайсы...')
        price = PriceList(price_path, log)
        log('Сверяю с прайсом...')
        if items:
            stats = compare(items, infos, price, log)
        if order_items:
            for it in order_items:
                it.setdefault('page', '')
                it.setdefault('no', '')
            compare(order_items, [dict(file=f, currency='') for f in {it['file'] for it in order_items}], price,
                    lambda *a: None)
    if progress:
        progress(0.9)
    os.makedirs(os.path.dirname(os.path.abspath(out_xlsx)), exist_ok=True)
    write_report(out_xlsx, infos, items, price, order_rows, order_items)
    if progress:
        progress(1.0)
    log(f'\nДокументов: {len(infos)}, позиций: {len(items)}')
    bad = sum(1 for it in items if it['check'] == 'проверить')
    if bad:
        log(f'  строк, где кол-во × цена ≠ сумма: {bad}')
    if order_items:
        log(f'Позиций в заказах: {len(order_items)}')
    log(f'Отчёт: {out_xlsx}')
    return dict(docs=len(infos), items=len(items), stats=stats, bad_lines=bad, errors=errors,
                orders=len(order_items), order_stats=ostats)


def main():
    args = sys.argv[1:]
    price, order = [], []
    for flag, dst in (('--price', price), ('--order', order)):
        while flag in args:
            i = args.index(flag)
            dst.append(args[i + 1])
            del args[i:i + 2]
    if len(args) < 1 or (len(args) < 2 and not order):
        print(__doc__)
        sys.exit(1)
    run(args[1:], args[0], price or None, order_paths=order or None)


if __name__ == '__main__':
    main()
