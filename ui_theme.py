"""Оформление Inventa в стиле сайта: бирюзовая шапка, светлый фон, скруглённые (squircle) карточки,
поля и кнопки, шрифт Manrope, название — рукописным Grand Hotel.

Скругления рисуются картинками (PIL, со сглаживанием) и подключаются как элементы темы ttk.
Шрифты лежат в папке fonts (лицензия SIL OFL) и подключаются только для этой программы.
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, font as tkfont

from PIL import Image, ImageDraw, ImageTk

# палитра сайта
INK = '#0f2f38'
INK_SOFT = '#46636b'
MUTED = '#7b9198'
TEAL = '#134e5e'
TEAL_2 = '#1d6a7c'
MIST = '#e3f1f4'
PAPER = '#f7fbfc'
CARD = '#ffffff'
LINE = '#d3e4e8'
FIELD = '#f3f8f9'
AMBER = '#f5a524'
AMBER_HOVER = '#eb9612'
AMBER_INK = '#3a2500'
OK = '#2f8f5b'
WARN = '#c9791a'
BAD = '#c0423a'

F = {}            # шрифты: F['body'], F['h1'] …
_IMAGES = []      # ссылки на картинки, чтобы их не удалил сборщик мусора
S = 1.0           # масштаб экрана (Windows 125 % -> 1.25)


def _base():
    return getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))


def load_fonts():
    """Подключает шрифты из папки fonts только для этой программы (Windows). До создания окна."""
    folder = os.path.join(_base(), 'fonts')
    if not os.path.isdir(folder) or not sys.platform.startswith('win'):
        return
    try:
        import ctypes
        for f in os.listdir(folder):
            if f.lower().endswith(('.ttf', '.otf')):
                ctypes.windll.gdi32.AddFontResourceExW(os.path.join(folder, f), 0x10, 0)   # FR_PRIVATE
    except Exception:
        pass


def _family(root, names):
    have = {f.lower() for f in tkfont.families(root)}
    for n in names:
        if n.lower() in have:
            return n
    return 'TkDefaultFont'


# ---------------------------------------------------------------- скругления
def _squircle(w, h, r, n=4.0, steps=24):
    """Контур прямоугольника со «сквиркл»-углами (суперэллипс), точки по часовой стрелке."""
    import math
    r = min(r, w / 2, h / 2)
    pts = []
    corners = ((w - r, r, 0), (w - r, h - r, 1), (r, h - r, 2), (r, r, 3))
    for cx, cy, q in corners:
        for i in range(steps + 1):
            t = (q - 1) * math.pi / 2 + i * (math.pi / 2) / steps
            c, s = math.cos(t), math.sin(t)
            x = cx + r * math.copysign(abs(c) ** (2 / n), c)
            y = cy + r * math.copysign(abs(s) ** (2 / n), s)
            pts.append((x, y))
    return pts


def round_image(w, h, r, fill, outer, border=None, bw=1, k=4):
    """Картинка w×h: скруглённая плашка цвета fill (с рамкой border) на фоне outer."""
    W, H = w * k, h * k
    img = Image.new('RGB', (W, H), outer)
    d = ImageDraw.Draw(img)
    if border:
        d.polygon(_squircle(W - 1, H - 1, r * k), fill=border)
        inset = bw * k
        pts = [(x + inset, y + inset) for x, y in _squircle(W - 1 - 2 * inset, H - 1 - 2 * inset, (r - bw) * k)]
        d.polygon(pts, fill=fill)
    else:
        d.polygon(_squircle(W - 1, H - 1, r * k), fill=fill)
    return img.resize((w, h), Image.LANCZOS)


def _photo(root, img):
    p = ImageTk.PhotoImage(img, master=root)
    _IMAGES.append(p)
    return p


def _px(v):
    return max(1, int(round(v * S)))


CENTER = 160


def _widen(img, b, cx, cy):
    """Раздвигает середину 9-частной картинки (b — ширина краёв): углы как есть, края и середина
    растягиваются. Края однородны вдоль растяжения, поэтому NEAREST — без потери качества."""
    w, h = img.size
    W, H = w + cx, h + cy
    out = Image.new(img.mode, (W, H))
    xs = ((0, b, 0, b), (b, w - b, b, W - b), (w - b, w, W - b, W))
    ys = ((0, b, 0, b), (b, h - b, b, H - b), (h - b, h, H - b, H))
    for sx0, sx1, dx0, dx1 in xs:
        for sy0, sy1, dy0, dy1 in ys:
            if sx1 > sx0 and sy1 > sy0 and dx1 > dx0 and dy1 > dy0:
                part = img.crop((sx0, sy0, sx1, sy1)).resize((dx1 - dx0, dy1 - dy0), Image.NEAREST)
                out.paste(part, (dx0, dy0))
    return out


def _elem(style, root, name, r, fill, outer, border=None, states=(), size=None):
    """Элемент темы из скруглённой картинки, растягиваемой по 9 частям."""
    sz = size or 2 * _px(r) + 8
    rr = _px(r)
    # Середину делаем большой: Tk заполняет растянутую середину копиями этого куска, и если он
    # крошечный, большая карточка рисуется тысячами мелких копий — на Windows это сотни миллисекунд.
    cx, cy = CENTER, (0 if size else CENTER)

    def img(f, b):
        return _photo(root, _widen(round_image(sz, sz, rr, f, outer, b), rr + 1, cx, cy))
    base = img(fill, border)
    st = [(spec, img(f, b)) for spec, f, b in states]
    style.element_create(name, 'image', base, *st, border=rr + 1, padding=0, sticky='nsew')


# ---------------------------------------------------------------- тема
def apply(root):
    global S
    S = max(1.0, root.winfo_fpixels('1i') / 96)
    sans = _family(root, ['Manrope', 'Segoe UI', 'Inter', 'Arial'])
    script = _family(root, ['Grand Hotel', 'Segoe Script', 'Brush Script MT'])
    mono = _family(root, ['Cascadia Mono', 'Consolas', 'JetBrains Mono', 'DejaVu Sans Mono', 'Courier New'])
    F.update(
        body=(sans, 10), small=(sans, 9), bold=(sans, 10, 'bold'),
        h1=(sans, 21, 'bold'), h2=(sans, 16, 'bold'), h3=(sans, 12, 'bold'),
        brand=(script, 22), eyebrow=(sans, 8, 'bold'), btn=(sans, 10, 'bold'),
        big=(sans, 11, 'bold'), mono=(mono, 9), icon=(sans, 16),
    )
    for name in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont'):
        try:
            tkfont.nametofont(name).configure(family=sans, size=10)
        except tk.TclError:
            pass
    root.configure(bg=PAPER)

    s = ttk.Style(root)
    s.theme_use('clam')
    s.configure('.', background=PAPER, foreground=INK, font=F['body'], bordercolor=LINE,
                lightcolor=PAPER, darkcolor=PAPER, troughcolor=MIST, focuscolor=TEAL_2,
                selectbackground=MIST, selectforeground=INK, insertcolor=INK)
    s.configure('TFrame', background=PAPER)
    s.configure('Flat.TFrame', background=CARD)
    s.configure('TLabel', background=PAPER, foreground=INK)
    s.configure('Muted.TLabel', background=PAPER, foreground=INK_SOFT)
    s.configure('H1.TLabel', background=PAPER, foreground=INK, font=F['h1'])
    s.configure('H2.TLabel', background=PAPER, foreground=INK, font=F['h2'])
    s.configure('Eyebrow.TLabel', background=PAPER, foreground=TEAL_2, font=F['eyebrow'])
    s.configure('Status.TLabel', background=PAPER, foreground=INK_SOFT, font=F['bold'])

    # карточки и поля
    _elem(s, root, 'Card.bg', 26, CARD, PAPER, LINE)
    _elem(s, root, 'CardHover.bg', 26, CARD, PAPER, TEAL_2)
    _elem(s, root, 'Field.bg', 18, FIELD, CARD, LINE)
    _elem(s, root, 'FieldP.bg', 18, FIELD, PAPER, LINE)
    for name, el, pad in (('Card.TFrame', 'Card.bg', _px(18)), ('CardHover.TFrame', 'CardHover.bg', _px(18)),
                          ('Field.TFrame', 'Field.bg', (_px(10), _px(6))),
                          ('FieldP.TFrame', 'FieldP.bg', (_px(12), _px(8)))):
        s.layout(name, [(el, {'sticky': 'nsew'})])
        s.configure(name, padding=pad, background=CARD if 'Field' not in name else FIELD)
    # промежуточные рамки для плавной подсветки карточки при наведении
    def _mix(a, b, t):
        a, b = a.lstrip('#'), b.lstrip('#')
        return '#%02x%02x%02x' % tuple(round(int(a[i:i + 2], 16) + (int(b[i:i + 2], 16) - int(a[i:i + 2], 16)) * t)
                                       for i in (0, 2, 4))
    for i in range(8):
        el = f'CardH{i}.bg'
        _elem(s, root, el, 26, CARD, PAPER, _mix(LINE, TEAL_2, i / 7))
        s.layout(f'CardH{i}.TFrame', [(el, {'sticky': 'nsew'})])
        s.configure(f'CardH{i}.TFrame', background=CARD)
    s.configure('Card.TFrame', background=CARD)
    s.configure('CardHover.TFrame', background=CARD)

    # кнопки-«пилюли»: (стиль, фон снаружи, цвет, при наведении, рамка, текст, отступы, шрифт, радиус)
    buttons = [
        ('TButton', CARD, CARD, MIST, LINE, TEAL, (_px(16), _px(9)), F['btn'], 17),
        ('Paper.TButton', PAPER, CARD, MIST, LINE, TEAL, (_px(18), _px(10)), F['btn'], 18),
        ('Accent.TButton', PAPER, AMBER, AMBER_HOVER, None, AMBER_INK, (_px(26), _px(12)), F['big'], 21),
        ('AccentCard.TButton', CARD, AMBER, AMBER_HOVER, None, AMBER_INK, (_px(24), _px(11)), F['big'], 20),
        ('Ghost.TButton', TEAL, TEAL, TEAL_2, '#3d7f8f', '#e3f1f4', (_px(16), _px(8)), F['small'], 15),
        ('Back.TButton', PAPER, PAPER, MIST, None, TEAL, (_px(12), _px(8)), F['btn'], 16),
    ]
    for name, outer, fill, hover, border, fg, pad, fnt, rad in buttons:
        el = name.replace('.', '_') + '.bg'
        _elem(s, root, el, rad, fill, outer, border,
              states=[('disabled', '#edf2f4', '#e1e9ec'), ('pressed', hover, border or hover),
                      ('active', hover, border or hover)])
        s.layout(name, [(el, {'sticky': 'nsew', 'children': [
            ('Button.padding', {'sticky': 'nsew', 'children': [('Button.label', {'sticky': 'nsew'})]})]})])
        s.configure(name, foreground=fg, padding=pad, font=fnt, anchor='center', background=outer)
        s.map(name, foreground=[('disabled', '#9fb2b8')])

    # поле ввода
    _elem(s, root, 'REntry.field', 16, FIELD, CARD, LINE, states=[('focus', FIELD, TEAL_2)])
    s.layout('TEntry', [('REntry.field', {'sticky': 'nsew', 'children': [
        ('Entry.padding', {'sticky': 'nsew', 'children': [('Entry.textarea', {'sticky': 'nsew'})]})]})])
    s.configure('TEntry', padding=(_px(14), _px(9)), foreground=INK, fieldbackground=FIELD)

    # полоса прогресса: скруглённые дорожка и заливка
    h = _px(10)
    _elem(s, root, 'Round.trough', h // 2, MIST, PAPER, size=h + 8)
    _elem(s, root, 'Round.pbar', h // 2, AMBER, MIST, size=h + 8)
    s.layout('Horizontal.TProgressbar', [('Round.trough', {'sticky': 'nsew', 'children': [
        ('Round.pbar', {'side': 'left', 'sticky': 'ns'})]})])
    s.configure('Horizontal.TProgressbar', thickness=h)
    return s


# ---------------------------------------------------------------- готовые детали
def card(parent, hover=False, **kw):
    """Скруглённая белая карточка (ttk.Frame)."""
    return ttk.Frame(parent, style='Card.TFrame', **kw)


def badge(parent, text, bg=TEAL, fg='#ffffff', size=24, parent_bg=CARD):
    """Кружок с номером шага."""
    size = _px(size)
    k = 4
    img = Image.new('RGB', (size * k, size * k), parent_bg)
    ImageDraw.Draw(img).ellipse((0, 0, size * k - 1, size * k - 1), fill=bg)
    p = _photo(parent, img.resize((size, size), Image.LANCZOS))
    return tk.Label(parent, image=p, text=text, compound='center', fg=fg, bg=parent_bg, font=F['bold'], bd=0)


def icon_tile(parent, text, size=46, parent_bg=CARD):
    """Скруглённая плашка с иконкой (на карточках задач)."""
    size = _px(size)
    p = _photo(parent, round_image(size, size, _px(17), MIST, parent_bg))
    return tk.Label(parent, image=p, text=text, compound='center', fg=TEAL, bg=parent_bg, font=F['icon'], bd=0)


def bullet(parent, text, bg=CARD):
    row = tk.Frame(parent, bg=bg)
    d = _px(7)
    k = 4
    img = Image.new('RGB', (d * k, d * k), bg)
    ImageDraw.Draw(img).ellipse((0, 0, d * k - 1, d * k - 1), fill=AMBER)
    p = _photo(parent, img.resize((d, d), Image.LANCZOS))
    tk.Label(row, image=p, bg=bg, bd=0).pack(side='left', padx=(2, 10))
    tk.Label(row, text=text, fg=INK_SOFT, bg=bg, font=F['body'], justify='left', anchor='w').pack(
        side='left', fill='x')
    return row


def listbox(parent, **kw):
    """-> (скруглённая рамка, список)."""
    fr = ttk.Frame(parent, style='Field.TFrame', padding=(_px(16), _px(10)))
    lb = tk.Listbox(fr, bg=FIELD, fg=INK, selectbackground='#cfe7ec', selectforeground=INK,
                    highlightthickness=0, relief='flat', bd=0, activestyle='none', font=F['body'], **kw)
    lb.pack(fill='both', expand=True)
    return fr, lb


def logo(root, size=28):
    """Иконка программы для шапки (app.png рядом с программой или внутри exe)."""
    try:
        img = Image.open(os.path.join(_base(), 'app.png')).convert('RGBA').resize((_px(size), _px(size)),
                                                                                 Image.LANCZOS)
        bg = Image.new('RGBA', img.size, TEAL)
        bg.alpha_composite(img)
        return _photo(root, bg.convert('RGB'))
    except Exception:
        return None
