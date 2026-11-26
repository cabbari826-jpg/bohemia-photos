"""Оформление Inventa в стиле сайта: бирюзовая шапка, светлый фон, белые карточки, янтарные кнопки.

Всё на стандартном tkinter/ttk (тема clam), без сторонних библиотек.
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, font as tkfont

# палитра сайта
INK = '#0f2f38'
INK_SOFT = '#46636b'
MUTED = '#7b9198'
TEAL = '#134e5e'
TEAL_2 = '#1d6a7c'
TEAL_DARK = '#0f3e4b'
MIST = '#e3f1f4'
PAPER = '#f7fbfc'
CARD = '#ffffff'
LINE = '#d3e4e8'
FIELD = '#f4f8f9'
AMBER = '#f5a524'
AMBER_HOVER = '#e8961a'
AMBER_INK = '#3a2500'
OK = '#2f8f5b'
WARN = '#c9791a'
BAD = '#c0423a'

F = {}          # шрифты: F['body'], F['h1'] …


def _family(root, names):
    have = {f.lower() for f in tkfont.families(root)}
    for n in names:
        if n.lower() in have:
            return n
    return 'TkDefaultFont'


def apply(root):
    sans = _family(root, ['Segoe UI Variable Text', 'Segoe UI', 'Inter', 'Helvetica Neue', 'Arial'])
    mono = _family(root, ['Cascadia Mono', 'Consolas', 'JetBrains Mono', 'DejaVu Sans Mono', 'Courier New'])
    F.update(
        body=(sans, 10), small=(sans, 9), bold=(sans, 10, 'bold'),
        h1=(sans, 20, 'bold'), h2=(sans, 15, 'bold'), h3=(sans, 12, 'bold'),
        brand=(sans, 14, 'bold'), eyebrow=(sans, 8, 'bold'), btn=(sans, 10, 'bold'),
        big=(sans, 11, 'bold'), mono=(mono, 9),
    )
    for name in ('TkDefaultFont', 'TkTextFont', 'TkMenuFont'):
        try:
            tkfont.nametofont(name).configure(family=sans, size=10)
        except tk.TclError:
            pass
    root.configure(bg=PAPER)
    root.option_add('*Listbox.font', F['body'])
    root.option_add('*TCombobox*Listbox.font', F['body'])

    s = ttk.Style(root)
    s.theme_use('clam')
    s.configure('.', background=PAPER, foreground=INK, font=F['body'], bordercolor=LINE,
                lightcolor=PAPER, darkcolor=PAPER, troughcolor=MIST, focuscolor=TEAL_2,
                selectbackground=MIST, selectforeground=INK, insertcolor=INK)
    s.configure('TFrame', background=PAPER)
    s.configure('Card.TFrame', background=CARD)
    s.configure('Header.TFrame', background=TEAL)
    s.configure('TLabel', background=PAPER, foreground=INK)
    s.configure('Card.TLabel', background=CARD, foreground=INK)
    s.configure('CardMuted.TLabel', background=CARD, foreground=INK_SOFT)
    s.configure('Muted.TLabel', background=PAPER, foreground=INK_SOFT)
    s.configure('H1.TLabel', background=PAPER, foreground=INK, font=F['h1'])
    s.configure('H2.TLabel', background=PAPER, foreground=INK, font=F['h2'])
    s.configure('CardH.TLabel', background=CARD, foreground=INK, font=F['h3'])
    s.configure('Step.TLabel', background=CARD, foreground=INK, font=F['bold'])
    s.configure('Eyebrow.TLabel', background=PAPER, foreground=TEAL_2, font=F['eyebrow'])
    s.configure('Status.TLabel', background=PAPER, foreground=INK_SOFT, font=F['bold'])

    def button(name, bg, fg, hover, border=None, pad=(16, 8), font=F['btn']):
        s.configure(name, background=bg, foreground=fg, bordercolor=border or bg, lightcolor=bg, darkcolor=bg,
                    padding=pad, font=font, relief='flat' if border is None else 'solid', borderwidth=1,
                    focusthickness=0, focuscolor=bg)
        s.map(name,
              background=[('disabled', '#e9eff1'), ('pressed', hover), ('active', hover)],
              lightcolor=[('disabled', '#e9eff1'), ('active', hover)],
              darkcolor=[('disabled', '#e9eff1'), ('active', hover)],
              bordercolor=[('disabled', '#dfe8eb'), ('active', border or hover)],
              foreground=[('disabled', '#9fb2b8')])

    button('TButton', CARD, TEAL, MIST, border=LINE, pad=(12, 6))
    button('Accent.TButton', AMBER, AMBER_INK, AMBER_HOVER, pad=(20, 9), font=F['big'])
    button('Teal.TButton', TEAL, '#ffffff', TEAL_2, pad=(16, 8))
    button('Ghost.TButton', TEAL, '#d9ecf0', TEAL_2, border='#2f6d7d', pad=(12, 5), font=F['small'])
    button('Back.TButton', PAPER, TEAL, MIST, pad=(8, 5))
    button('Link.TButton', CARD, TEAL_2, MIST, pad=(10, 6))

    s.configure('TEntry', fieldbackground=FIELD, foreground=INK, bordercolor=LINE, lightcolor=FIELD,
                darkcolor=FIELD, padding=(8, 6))
    s.map('TEntry', bordercolor=[('focus', TEAL_2)], lightcolor=[('focus', FIELD)])
    s.configure('Horizontal.TProgressbar', troughcolor=MIST, background=AMBER, bordercolor=MIST,
                lightcolor=AMBER, darkcolor=AMBER, thickness=8)
    s.configure('Vertical.TScrollbar', background=MIST, troughcolor=FIELD, bordercolor=FIELD,
                arrowcolor=INK_SOFT, lightcolor=MIST, darkcolor=MIST)
    return s


def card(parent, **kw):
    """Белая карточка с тонкой рамкой."""
    return tk.Frame(parent, bg=CARD, highlightbackground=LINE, highlightcolor=LINE, highlightthickness=1, **kw)


def badge(parent, text, bg=TEAL, fg='#ffffff', size=24, parent_bg=CARD):
    """Кружок с номером шага."""
    c = tk.Canvas(parent, width=size, height=size, bg=parent_bg, highlightthickness=0, bd=0)
    c.create_oval(1, 1, size - 1, size - 1, fill=bg, outline=bg)
    c.create_text(size / 2, size / 2, text=text, fill=fg, font=F['bold'])
    return c


def bullet(parent, text, bg=CARD):
    row = tk.Frame(parent, bg=bg)
    tk.Label(row, text='◆', fg=AMBER, bg=bg, font=F['small']).pack(side='left', anchor='n', padx=(0, 8))
    tk.Label(row, text=text, fg=INK_SOFT, bg=bg, font=F['body'], justify='left', anchor='w').pack(
        side='left', fill='x')
    return row


def listbox(parent, **kw):
    """-> (рамка, список): у Listbox нет внутренних отступов, поэтому он лежит в рамке с полями."""
    fr = tk.Frame(parent, bg=FIELD, highlightthickness=1, highlightbackground=LINE, highlightcolor=TEAL_2,
                  padx=8, pady=5)
    lb = tk.Listbox(fr, bg=FIELD, fg=INK, selectbackground='#cfe7ec', selectforeground=INK,
                    highlightthickness=0, relief='flat', bd=0, activestyle='none', font=F['body'], **kw)
    lb.pack(fill='both', expand=True)
    return fr, lb


def logo(root, size=28):
    """Иконка программы для шапки (app.png рядом с программой или внутри exe)."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    p = os.path.join(base, 'app.png')
    try:
        from PIL import Image, ImageTk
        img = Image.open(p).convert('RGBA').resize((size, size), Image.LANCZOS)
        return ImageTk.PhotoImage(img, master=root)
    except Exception:
        return None
