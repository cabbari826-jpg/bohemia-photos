"""«Живые» элементы интерфейса в духе Material 3 Expressive на обычном tkinter.

* Spring — пружинная анимация (с лёгким перелётом и возвратом).
* MotionButton — кнопка-«таблетка»: плавно подсвечивается при наведении; при нажатии сжимается,
  углы «перетекают» из круглых в более квадратные, при отпускании пружинит обратно.
* Field — скруглённое поле со списком или строкой ввода внутри: при фокусе рамка плавно темнеет,
  при перетаскивании файлов над полем — подсвечивается и «вдыхает», при сбросе — подпрыгивает.
* «Вес» файла: чем тяжелее сброшенный файл, тем сильнее поле проседает и пружинит, раздвигая
  соседние элементы, будто на него что-то упало; потом всё возвращается на место.
* flash_items — новые строки списка коротко вспыхивают янтарным и гаснут.

Картинки форм рисуются через PIL со сглаживанием и кешируются, поэтому анимация лёгкая (~60 кадров/с).
"""
import time
import tkinter as tk

from PIL import Image, ImageDraw, ImageTk

import ui_theme as ui

FRAME_MS = 8          # ~120 запросов кадра в секунду; реальный шаг считается по часам


def _rgb(c):
    c = c.lstrip('#')
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t):
    t = max(0.0, min(1.0, t))
    a, b = _rgb(a), _rgb(b)
    return '#%02x%02x%02x' % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


class Spring:
    """Физическая пружина, считается по реальному времени (а не по кадрам), поэтому движение плавное,
    даже если таймер Windows срабатывает неровно. freq — частота колебаний (Гц, больше — быстрее),
    zeta — затухание (1 — без перелёта, меньше — с отскоком)."""

    def __init__(self, widget, on_step, value=0.0, freq=4.0, zeta=0.8):
        import math
        self.w, self.on_step = widget, on_step
        self.value, self.target, self.vel = value, value, 0.0
        self.omega, self.zeta = 2 * math.pi * freq, zeta
        self._job = None
        self._t = None

    def to(self, target, kick=0.0):
        """kick — начальная скорость (в единицах значения за секунду)."""
        self.target = target
        self.vel += kick
        if self._job is None:
            self._t = time.perf_counter()
            self._job = self.w.after(FRAME_MS, self._tick)

    def advance(self, dt):
        """Шаг физики на dt секунд (мелкими подшагами — устойчиво при любом dt). -> True, если ещё движется."""
        w, z = self.omega, self.zeta
        n = max(1, int(dt / 0.002))
        h = dt / n
        for _ in range(n):
            a = -w * w * (self.value - self.target) - 2 * z * w * self.vel
            self.vel += a * h
            self.value += self.vel * h
        if abs(self.vel) < 0.01 and abs(self.target - self.value) < 0.002:
            self.value, self.vel = self.target, 0.0
            return False
        return True

    def _tick(self):
        self._job = None
        t = time.perf_counter()
        dt = min(0.05, max(0.001, t - (self._t or t)))
        self._t = t
        moving = self.advance(dt)
        if moving:
            self._job = self.w.after(FRAME_MS, self._tick)
        try:
            self.on_step()
        except tk.TclError:
            pass


_cache = {}


def shape(w, h, r, fill, outer, border=None, bw=1.0, inset=0.0, k=3):
    """Скруглённая плашка w×h (с отступом inset от краёв) на фоне outer; результат кешируется."""
    key = (w, h, round(r, 2), fill, outer, border, round(bw, 2), round(inset, 2))
    img = _cache.get(key)
    if img is None:
        W, H = w * k, h * k
        im = Image.new('RGB', (W, H), outer)
        d = ImageDraw.Draw(im)
        x0, y0, x1, y1 = inset * k, inset * k, W - 1 - inset * k, H - 1 - inset * k
        rr = max(1, min(r * k, (x1 - x0) / 2, (y1 - y0) / 2))
        if border:
            d.rounded_rectangle((x0, y0, x1, y1), rr, fill=border)
            b = bw * k
            d.rounded_rectangle((x0 + b, y0 + b, x1 - b, y1 - b), max(1, rr - b), fill=fill)
        else:
            d.rounded_rectangle((x0, y0, x1, y1), rr, fill=fill)
        img = im.resize((w, h), Image.LANCZOS)
        if len(_cache) > 3000:
            _cache.clear()
        _cache[key] = img
    return img


# ------------------------------------------------------------------ кнопка
KINDS = {
    #            fill        hover       border     text        font
    'secondary': (ui.CARD, ui.MIST, ui.LINE, ui.TEAL, 'btn'),
    'accent': (ui.AMBER, ui.AMBER_HOVER, None, ui.AMBER_INK, 'big'),
    'ghost': (ui.TEAL, ui.TEAL_2, '#3d7f8f', '#e3f1f4', 'small'),
    'back': (ui.PAPER, ui.MIST, None, ui.TEAL, 'btn'),
}


class MotionButton(tk.Canvas):
    def __init__(self, master, text='', command=None, kind='secondary', width=132, height=40, bg=ui.CARD,
                 state='normal', **kw):
        self.W, self.H = ui._px(width), ui._px(height)
        m = ui._px(3)
        super().__init__(master, width=self.W + 2 * m, height=self.H + 2 * m, bg=bg, highlightthickness=0, bd=0,
                         cursor='hand2', **kw)
        self.m, self.outer, self.kind, self.command = m, bg, kind, command
        self.state = state
        self.hover = Spring(self, self._draw, freq=3.5, zeta=1.0)
        self.press = Spring(self, self._draw, freq=5.0, zeta=0.5)
        self._img = self.create_image(0, 0, anchor='nw')
        self._txt = self.create_text((self.W + 2 * m) / 2, (self.H + 2 * m) / 2, text=text,
                                     font=ui.F[KINDS[kind][4]])
        self.bind('<Enter>', lambda e: self.hover.to(1))
        self.bind('<Leave>', lambda e: (self.hover.to(0), self.press.to(0)))
        self.bind('<ButtonPress-1>', self._down)
        self.bind('<ButtonRelease-1>', self._up)
        self._photo = None
        self._draw()

    def _down(self, e):
        if self.state != 'disabled':
            self.press.to(1)

    def _up(self, e):
        if self.state == 'disabled':
            return
        inside = 0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height()
        self.press.to(0, kick=-3.0)           # отпустили — пружинит, чуть перелетая
        if inside and self.command:
            self.after(60, self.command)

    def _draw(self):
        fill, hov, border, fg, _ = KINDS[self.kind]
        p, h = self.press.value, self.hover.value
        if self.state == 'disabled':
            fill = hov = '#edf2f4'
            border = '#e1e9ec' if border else None
            fg = '#9fb2b8'
            p = h = 0
        c = mix(fill, hov, h + 0.5 * max(0, p))
        b = mix(border, ui.TEAL_2, 0.5 * max(0, p)) if border else None
        r = self.H / 2 * (1 - 0.42 * max(0.0, min(1.2, p)))      # круглые углы -> «сквиркл» при нажатии
        inset = self.m + ui._px(2.2) * p                          # сжатие (при перелёте — чуть больше обычного)
        inset = max(0.0, inset)
        img = shape(self.W + 2 * self.m, self.H + 2 * self.m, r, c, self.outer, b, inset=round(inset * 4) / 4)
        self._photo = ImageTk.PhotoImage(img, master=self)
        self.itemconfigure(self._img, image=self._photo)
        self.itemconfigure(self._txt, fill=fg)

    def configure(self, cnf=None, **kw):
        cnf = dict(cnf or {}, **kw)
        redraw = False
        if 'text' in cnf:
            self.itemconfigure(self._txt, text=cnf.pop('text'))
        if 'state' in cnf:
            self.state = str(cnf.pop('state'))
            super().configure(cursor='arrow' if self.state == 'disabled' else 'hand2')
            redraw = True
        if 'command' in cnf:
            self.command = cnf.pop('command')
        if cnf:
            super().configure(**cnf)
        if redraw:
            self._draw()

    config = configure

    def cget(self, key):
        if key == 'state':
            return self.state
        if key == 'text':
            return self.itemcget(self._txt, 'text')
        return super().cget(key)

    def invoke(self):
        if self.state != 'disabled' and self.command:
            self.command()


# ------------------------------------------------------------------ поле
class Field(tk.Canvas):
    """Скруглённое поле; внутри — виджет, созданный make(self) (Listbox или Entry)."""

    def __init__(self, master, make, height=40, bg=ui.CARD, pad=(14, 8), **kw):
        super().__init__(master, height=ui._px(height), bg=bg, highlightthickness=0, bd=0, **kw)
        self.base_h = ui._px(height)
        self.outer, self.pad = bg, (ui._px(pad[0]), ui._px(pad[1]))
        self.impact = Spring(self, self._bump, freq=2.4, zeta=0.32)
        self.amp = 0
        self.m = ui._px(4)                            # запас под «вдох» при перетаскивании
        self.focus = Spring(self, self._draw, freq=3.5, zeta=1.0)
        self.drag = Spring(self, self._draw, freq=3.0, zeta=0.75)
        self.pop = Spring(self, self._draw, freq=3.2, zeta=0.4)
        self._img = self.create_image(0, 0, anchor='nw')
        self.inner = make(self)
        self._win = self.create_window(0, 0, anchor='nw', window=self.inner)
        self.inner.bind('<FocusIn>', lambda e: self.focus.to(1), add='+')
        self.inner.bind('<FocusOut>', lambda e: self.focus.to(0), add='+')
        self.bind('<Configure>', lambda e: self._draw())
        self._photo = None

    def drag_over(self, on):
        self.drag.to(1 if on else 0)

    def dropped(self, size=0):
        """Сброс файлов: поле «подпрыгивает». size — сколько весят файлы (байт): тяжёлые бьют сильнее —
        поле проседает (растёт вниз и сдвигает всё, что ниже) и пружинит обратно."""
        import math
        self.drag.to(0)
        w = min(1.0, max(0.15, math.log10(max(size, 1) / 20_000 + 1) / 2.3))   # 0.15 … 1 (≈ от 10 КБ до 4 МБ+)
        self.pop.value = 0.4 + 0.6 * w
        self.pop.to(0)
        self.amp = ui._px(26) * w
        self.impact.value = 0.0
        self.impact.to(0, kick=23.0)        # толчок вниз, перелёт вверх, затухание

    def _bump(self):
        extra = self.amp * self.impact.value
        super().configure(height=max(10, round(self.base_h + extra)))

    def _draw(self):
        w, h = self.winfo_width(), self.winfo_height()
        if w < 10 or h < 10:
            return
        f, d, p = self.focus.value, self.drag.value, self.pop.value
        fill = mix(ui.FIELD, '#e4f3f1', d)
        border = mix(mix(ui.LINE, ui.TEAL_2, f), ui.TEAL, d)
        bw = 1 + 1.2 * max(f * 0.6, d)
        inset = self.m * (1 - d) + ui._px(3) * p                 # перетаскивание — поле «вдыхает»
        inset = max(0.0, inset)
        img = shape(w, h, ui._px(16), fill, self.outer, border, bw=round(bw * 4) / 4, inset=round(inset * 4) / 4)
        self._photo = ImageTk.PhotoImage(img, master=self)
        self.itemconfigure(self._img, image=self._photo)
        px, py = self.pad
        self.coords(self._win, self.m + px, self.m + py)
        self.itemconfigure(self._win, width=max(10, w - 2 * (self.m + px)), height=max(10, h - 2 * (self.m + py)))
        try:
            self.inner.configure(bg=fill)
        except tk.TclError:
            pass
        try:
            self.inner.configure(readonlybackground=fill)
        except tk.TclError:
            pass


def entry_field(master, var, height=44, bg=ui.CARD):
    def make(parent):
        return tk.Entry(parent, textvariable=var, bg=ui.FIELD, fg=ui.INK, relief='flat', bd=0,
                        highlightthickness=0, insertbackground=ui.INK, font=ui.F['body'],
                        selectbackground='#cfe7ec', selectforeground=ui.INK)
    return Field(master, make, height=height, bg=bg, pad=(14, 11))


def list_field(master, rows=2, bg=ui.CARD, **kw):
    line = ui._px(22)

    def make(parent):
        return tk.Listbox(parent, bg=ui.FIELD, fg=ui.INK, selectbackground='#cfe7ec', selectforeground=ui.INK,
                          highlightthickness=0, relief='flat', bd=0, activestyle='none', font=ui.F['body'], **kw)
    return Field(master, make, height=rows * line + 2 * ui._px(10) + 8, bg=bg, pad=(14, 8))


def flash_items(lb, indices, steps=14, ms=40):
    """Новые строки списка вспыхивают янтарным и плавно гаснут."""
    def run(i):
        t = i / steps
        base = lb.cget('bg')
        c = mix('#fde7b8', base, t)
        for j in indices:
            try:
                lb.itemconfigure(j, background=c)
            except tk.TclError:
                return
        if i < steps:
            lb.after(ms, run, i + 1)
        else:
            for j in indices:
                try:
                    lb.itemconfigure(j, background='')
                except tk.TclError:
                    pass
    run(0)


class CardHover:
    """Карточка на главном экране: рамка плавно меняет цвет при наведении (набор стилей рамки)."""
    N = 8

    def __init__(self, frame, styles):
        self.f, self.styles = frame, styles
        self.s = Spring(frame, self._draw, freq=3.0, zeta=1.0)

    def set(self, on):
        self.s.to(1 if on else 0)

    def _draw(self):
        i = max(0, min(len(self.styles) - 1, round(self.s.value * (len(self.styles) - 1))))
        self.f.configure(style=self.styles[i])
