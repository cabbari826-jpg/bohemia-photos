"""Наши заказы поставщикам (Excel) и сверка с их Proforma / Invoice.

Заказ — Excel в формате, в каком его отправляют поставщику: Reference / Description / Price / Quantity / Total,
бланки Le Creuset (много строк с нулевым количеством, несколько листов), Bormioli (кол-во в Pallets / Layers /
Cartons / Sales Units), Bohemia (QTY in PCS и QTY in Sets) и т.п. Берём строки с количеством > 0.

Сверка: по коду (без учёта разделителей: 021-004-90-000-0 = 021-004-90-000/0) сравниваем заказ с документами
поставщика: чего нет в документах, что лишнее, где другое количество или цена.
"""
import os
import re
import collections

import bohemia_photos as core
from invoice_reader import role_of, parse_num

# колонка «кол-во», «цена», «сумма» — если таких несколько, какую брать
PREFER = {
    'qty': (r'pcs|pieces|шт|order|requested|sales\s*unit|quantity|кол', r'set|pack|carton|box|layer|pallet|остат|stock'),
    'price': (r'1\s*pc|per\s*pc|unit|item|за\s*шт|price', r'set|srrp|rrp|retail|розн'),
    'total': (r'pc|total|amount|сумма', r'set'),
}
QTY_EXTRA = re.compile(r'quan|кол-?во|количество|sales\s*units?', re.I)


def _role(text):
    t = str(text).strip()
    r = role_of(t)
    if r is None and QTY_EXTRA.search(t) and len(t) < 40:
        return 'qty'
    return r


def _pick(cols, role, titles):
    """Из нескольких колонок одной роли выбираем лучшую по названию."""
    good, bad = PREFER.get(role, ('', '^$'))
    def score(c):
        t = titles[c].lower()
        return (1 if re.search(bad, t) else 0, 0 if re.search(good, t) else 1, c)
    return min(cols, key=score)


def _sheets(path):
    if path.lower().endswith(('.xlsx', '.xlsm')):
        import openpyxl
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            return [(ws.title, [tuple(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
        except Exception:
            pass
    rows, _ = core.read_price(path)
    return [('', rows)]


def _read_sheet(rows, name, sheet):
    head = None
    for i, row in enumerate(rows[:60]):
        roles = collections.defaultdict(list)
        titles = {}
        for c, v in enumerate(row):
            if v in (None, ''):
                continue
            r = _role(v)
            if r:
                roles[r].append(c)
                titles[c] = str(v)
        if (roles.get('code') or roles.get('desc')) and (roles.get('qty') or roles.get('total')) and len(roles) >= 3:
            head = (i, roles, titles)
            break
    if not head:
        return []
    i, roles, titles = head
    start = i + 1
    # подзаголовок под «Requested quantity»: Pallets / Layers / Cartons / Sales Units
    if start < len(rows):
        sub = rows[start]
        subt = {c: str(v) for c, v in enumerate(sub) if v not in (None, '')}
        if subt and all(parse_num(v) is None for v in subt.values()):
            q = [c for c, t in subt.items() if re.search(r'sales\s*unit|pcs|pieces|шт', t, re.I)] or \
                [c for c, t in subt.items() if re.search(r'carton|box', t, re.I)]
            if q:
                roles['qty'] = [q[0]]
                titles[q[0]] = subt[q[0]]
            start += 1
    col = {}
    for r, cs in roles.items():
        cs = [c for c in cs if not (r == 'code' and re.search(r'ean|bar', titles[c], re.I))] or cs
        col[r] = _pick(cs, r, titles) if len(cs) > 1 else cs[0]
    items = []
    for rn, row in enumerate(rows[start:], start + 1):
        def g(r):
            c = col.get(r)
            return row[c] if c is not None and c < len(row) else None
        code = g('code')
        if isinstance(code, float) and code.is_integer():
            code = int(code)
        code = str(code).strip() if code not in (None, '') else ''
        desc = re.sub(r'\s+', ' ', str(g('desc') or '')).strip()
        qty, price, total = (parse_num(g(r)) for r in ('qty', 'price', 'total'))
        if not code or not re.search(r'\d', code) or qty is None or qty <= 0:
            continue
        if total is None and price is not None:
            total = round(qty * price, 2)
        items.append(dict(file=name, sheet=sheet, row=rn, code=code, desc=desc, qty=qty, price=price,
                          total=round(total, 2) if total is not None else None))
    return items


def read_order(path, log=print):
    """-> список позиций заказа (кол-во > 0). Листы: первый с позициями; следующие — только если коды
    в них другие (бланк на нескольких листах), а не тот же заказ ещё раз (выгрузка из 1С и т.п.)."""
    name = os.path.basename(path)
    out, seen = [], set()
    for sheet, rows in _sheets(path):
        if re.search(r'terms|condition|услови', sheet, re.I):
            continue
        items = _read_sheet(rows, name, sheet)
        if not items:
            continue
        keys = {core.code_key(it['code']) for it in items}
        if out and len(keys & seen) > 0.3 * len(keys):
            continue
        out += items
        seen |= keys
    log(f'  заказ {name}: позиций {len(out)}')
    if not out:
        log('    ! не нашёл позиций с количеством (нужны колонки Code/Reference и Quantity)')
    return out


# ---------------------------------------------------------------- сверка
def _keys(code):
    c = core.clean_code(code)
    ks = [core.code_key(c)]
    sc = core.short_code(c)
    if sc:
        ks.append(core.code_key(sc))
    return [k for k in ks if k]


def _group(items):
    """Позиции одного кода складываем (один код бывает в нескольких строках или документах)."""
    g = collections.OrderedDict()
    for it in items:
        ks = _keys(it['code'])
        if not ks:
            continue
        k = ks[0]
        e = g.setdefault(k, dict(code=it['code'], desc=it['desc'], qty=0.0, prices=[], files=[], keys=ks, total=0.0))
        e['qty'] += it['qty'] or 0
        e['total'] += it['total'] or 0
        if it.get('price') is not None:
            e['prices'].append(it['price'])
        if it['file'] not in e['files']:
            e['files'].append(it['file'])
        if not e['desc']:
            e['desc'] = it['desc']
    return g


def _price(e):
    if not e['prices']:
        return None
    # средняя по сумме, если цены в строках разные
    if len(set(round(p, 4) for p in e['prices'])) == 1:
        return e['prices'][0]
    return round(e['total'] / e['qty'], 4) if e['qty'] else e['prices'][0]


def match_key(k, index, keys_sorted):
    if k in index:
        return k
    # один код — начало другого (код с цветом/упаковкой у одной из сторон), не короче 6 знаков
    import bisect
    if len(k) >= 6:
        j = bisect.bisect_left(keys_sorted, k)
        if j < len(keys_sorted) and keys_sorted[j].startswith(k) and len(keys_sorted[j]) - len(k) <= 4:
            return keys_sorted[j]
        for n in range(len(k) - 1, max(5, len(k) - 5), -1):
            if k[:n] in index:
                return k[:n]
    return None


def _pack_ratio(oq, dq, op, dp):
    """Заказ в штуках, документ в наборах (или наоборот): кол-во отличается в N раз, цена — в N раз обратно."""
    if not (oq and dq and op and dp) or abs(oq - dq) < 1e-6:
        return None
    r = oq / dq
    if r < 1:
        r = 1 / r
        if abs(dp * r - op) <= max(0.02, op * 0.003) * r and abs(r - round(r)) < 1e-6 and round(r) > 1:
            return int(round(r))
        return None
    if abs(r - round(r)) < 1e-6 and round(r) > 1 and abs(op * r - dp) <= max(0.02, dp * 0.003):
        return int(round(r))
    return None


def compare_orders(order_items, doc_items, tol=0.011):
    """-> строки сверки: dict(code, desc, o_qty, o_price, o_sum, o_files, d_qty, d_price, d_sum, d_files, status)."""
    orders, docs = _group(order_items), _group(doc_items)
    alias = {}
    for k, e in docs.items():
        for a in e['keys']:
            alias.setdefault(a, k)
    keys_sorted = sorted(alias)
    used = set()
    out = []
    for k, o in orders.items():
        dk = None
        for a in o['keys']:
            m = match_key(a, alias, keys_sorted)
            if m and alias[m] not in used:
                dk = alias[m]
                break
        d = docs.get(dk) if dk else None
        row = dict(code=o['code'], desc=o['desc'], o_qty=o['qty'], o_price=_price(o), o_sum=round(o['total'], 2),
                   o_files=', '.join(o['files']), d_code='', d_qty=None, d_price=None, d_sum=None, d_files='')
        if d is None:
            row['status'] = 'нет в документах поставщика'
        else:
            used.add(dk)
            row.update(d_code=d['code'], d_qty=d['qty'], d_price=_price(d), d_sum=round(d['total'], 2),
                       d_files=', '.join(d['files']))
            st = []
            op, dp = row['o_price'], row['d_price']
            pack = _pack_ratio(o['qty'], d['qty'], op, dp)
            if pack:
                # в документе наборами/коробками: 8 × 3.67 = 48 × 0.6118
                row['status'] = f'совпадает (в документе упаковками по {pack} шт)'
                out.append(row)
                continue
            if abs(d['qty'] - o['qty']) > 1e-6:
                st.append('кол-во меньше заказа' if d['qty'] < o['qty'] else 'кол-во больше заказа')
            if op is not None and dp is not None and abs(dp - op) > max(tol, abs(op) * 0.001):
                st.append('цена выше заказа' if dp > op else 'цена ниже заказа')
            row['status'] = ', '.join(st) or 'совпадает'
        out.append(row)
    # код чуть отличается (последние знаки — упаковка/версия), а цена та же: считаем одной позицией
    for row in out:
        if row['status'] != 'нет в документах поставщика':
            continue
        ok = core.code_key(row['code'])
        best = None
        for k, d in docs.items():
            if k in used:
                continue
            n = len(os.path.commonprefix([ok, k]))
            dp = _price(d)
            if n >= max(6, len(ok) - 3, len(k) - 3) and row['o_price'] is not None and dp is not None and \
                    abs(dp - row['o_price']) <= max(0.011, row['o_price'] * 0.002):
                if best is None or n > best[0]:
                    best = (n, k)
        if best:
            d = docs[best[1]]
            used.add(best[1])
            row.update(d_code=d['code'], d_qty=d['qty'], d_price=_price(d), d_sum=round(d['total'], 2),
                       d_files=', '.join(d['files']))
            st = ['код отличается']
            if abs(d['qty'] - row['o_qty']) > 1e-6:
                st.append('кол-во меньше заказа' if d['qty'] < row['o_qty'] else 'кол-во больше заказа')
            row['status'] = ', '.join(st)
    for k, d in docs.items():
        if k in used:
            continue
        out.append(dict(code=d['code'], desc=d['desc'], o_qty=None, o_price=None, o_sum=None, o_files='',
                        d_code=d['code'], d_qty=d['qty'], d_price=_price(d), d_sum=round(d['total'], 2),
                        d_files=', '.join(d['files']), status='нет в заказе (лишнее)'))
    return out


ORDER_FILLS = {'совпадает': '#C6EFCE', 'код отличается': '#DDEBF7', 'нет в документах поставщика': '#FFC7CE', 'нет в заказе (лишнее)': '#FFC7CE'}


def fill_of(status):
    if status.startswith('совпадает'):
        return ORDER_FILLS['совпадает']
    if status in ORDER_FILLS:
        return ORDER_FILLS[status]
    return '#FFEB9C'
