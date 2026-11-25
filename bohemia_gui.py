#!/usr/bin/env python3
"""Окно программы. Сначала выбирается действие:
  • «Proforma / Invoice → Excel» — PDF или Excel документов в Excel-отчёт, по желанию сверка с прайсами;
  • «Фото из каталогов» — вырезка фото по прайсам (бренд каждого прайса — Bohemia или Guzzini —
    определяется сам, его можно поменять).
"""
import os
import sys
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import bohemia_photos as core
import guzzini_photos
import invoice_reader

APP_TITLE = 'Bohemia Photos — каталоги и документы'
BRANDS = {'bohemia': ('Bohemia', core), 'guzzini': ('Guzzini', guzzini_photos)}
PRICE_TYPES = [('Excel', '*.xlsx *.xls *.csv'), ('Все файлы', '*.*')]


def open_path(p):
    if not p or not os.path.exists(p):
        return
    if sys.platform.startswith('win'):
        os.startfile(p)
    elif sys.platform == 'darwin':
        os.system(f'open "{p}"')
    else:
        os.system(f'xdg-open "{p}"')


def path_row(parent, r, label, var, cmd):
    ttk.Label(parent, text=label).grid(row=r, column=0, sticky='w', pady=4)
    ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky='ew', padx=6)
    ttk.Button(parent, text='Обзор…', command=cmd).grid(row=r, column=2, sticky='ew')


class FileList(ttk.Frame):
    """Список файлов с кнопками «Добавить…» / «Убрать». label(path) — как показывать строку."""

    def __init__(self, master, title, filetypes, label=os.path.basename, extra=None, height=4):
        super().__init__(master)
        self.paths, self.title, self.filetypes, self.label = [], title, filetypes, label
        self.columnconfigure(0, weight=1)
        self.box = tk.Listbox(self, height=height, selectmode='extended', activestyle='none')
        self.box.grid(row=0, column=0, sticky='ew')
        btns = ttk.Frame(self)
        btns.grid(row=0, column=1, sticky='n', padx=(6, 0))
        ttk.Button(btns, text='Добавить…', command=self.add).pack(fill='x')
        ttk.Button(btns, text='Убрать', command=self.remove).pack(fill='x', pady=2)
        for text, cmd in (extra or []):
            ttk.Button(btns, text=text, command=cmd).pack(fill='x')
        self.on_add = None

    def add(self):
        ps = filedialog.askopenfilenames(title=self.title, filetypes=self.filetypes)
        for p in ps:
            if p not in self.paths:
                self.paths.append(p)
                self.box.insert('end', self.label(p))
        if ps and self.on_add:
            self.on_add(ps)

    def remove(self):
        for i in reversed(self.box.curselection()):
            self.box.delete(i)
            del self.paths[i]

    def refresh(self):
        sel = self.box.curselection()
        self.box.delete(0, 'end')
        for p in self.paths:
            self.box.insert('end', self.label(p))
        for i in sel:
            self.box.selection_set(i)


class DocsScreen(ttk.Frame):
    """Proforma / Invoice -> Excel."""
    action = 'Сделать Excel'
    heading = 'Proforma / Invoice → Excel'

    def __init__(self, master):
        super().__init__(master)
        self.out = tk.StringVar()
        frm = ttk.Frame(self)
        frm.pack(fill='x')
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text='1. Документы (PDF или Excel):').grid(row=0, column=0, sticky='nw', pady=4)
        self.docs = FileList(frm, 'Proforma / Invoice',
                             [('PDF и Excel', '*.pdf *.PDF *.xlsx *.xls *.csv'), ('Все файлы', '*.*')])
        self.docs.grid(row=0, column=1, columnspan=2, sticky='ew', padx=6)
        self.docs.on_add = self._suggest_out
        ttk.Label(frm, text='2. Прайсы для сверки\n(необязательно, можно несколько):').grid(
            row=1, column=0, sticky='nw', pady=4)
        self.prices = FileList(frm, 'Прайс-листы', PRICE_TYPES, height=3)
        self.prices.grid(row=1, column=1, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Label(frm, text='3. Куда сохранить Excel:').grid(row=2, column=0, sticky='w', pady=4)
        ttk.Entry(frm, textvariable=self.out).grid(row=2, column=1, sticky='ew', padx=6)
        ttk.Button(frm, text='Обзор…', command=self.pick_out).grid(row=2, column=2, sticky='ew')

    def _suggest_out(self, ps):
        if not self.out.get():
            self.out.set(os.path.join(os.path.dirname(ps[0]), 'documents_report.xlsx'))

    def pick_out(self):
        p = filedialog.asksaveasfilename(title='Куда сохранить Excel', defaultextension='.xlsx',
                                         filetypes=[('Excel', '*.xlsx')], initialfile='documents_report.xlsx')
        if p:
            self.out.set(p)

    def result_path(self):
        return self.out.get().strip()

    def job(self):
        files, prices, out = list(self.docs.paths), list(self.prices.paths), self.out.get().strip()
        if not files:
            messagebox.showwarning(APP_TITLE, 'Нажмите «Добавить…» и выберите PDF или Excel файлы Proforma / Invoice.')
            return None
        if not out:
            messagebox.showwarning(APP_TITLE, 'Укажите, куда сохранить Excel.')
            return None
        if not out.lower().endswith('.xlsx'):
            out += '.xlsx'
            self.out.set(out)

        def work(log, progress):
            r = invoice_reader.run(files, out, prices or None, log=log, progress=progress)
            msg = f'Готово. Документов: {r["docs"]}, позиций: {r["items"]}\n'
            if r['bad_lines']:
                msg += f'Строк, где кол-во × цена ≠ сумма: {r["bad_lines"]}\n'
            if r['stats']:
                msg += '\nСверка с прайсом:\n' + '\n'.join(f'  {k}: {v}' for k, v in r['stats'].most_common())
            if r['errors']:
                msg += '\n\nНе прочитаны:\n' + '\n'.join(r['errors'])
            return msg + f'\n\nExcel: {out}', out
        return work


class PhotosScreen(ttk.Frame):
    """Фото из каталогов по одному или нескольким прайсам."""
    action = 'Вырезать фото'
    heading = 'Фото из каталогов'

    def __init__(self, master):
        super().__init__(master)
        self.brand = {}
        self.catalogs, self.existing, self.out = tk.StringVar(), tk.StringVar(), tk.StringVar()
        frm = ttk.Frame(self)
        frm.pack(fill='x')
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text='1. Прайс-листы\n(можно несколько брендов):').grid(row=0, column=0, sticky='nw', pady=4)
        self.prices = FileList(frm, 'Прайс-листы', PRICE_TYPES, label=self._label,
                               extra=[('Сменить бренд', self.toggle_brand)])
        self.prices.grid(row=0, column=1, columnspan=2, sticky='ew', padx=6)
        self.prices.on_add = self._added
        path_row(frm, 1, '2. Папка с каталогами (PDF):', self.catalogs, self.pick_catalogs)
        path_row(frm, 2, '3. Имеющиеся фото (необязательно):', self.existing, self.pick_existing)
        path_row(frm, 3, '4. Куда сохранить результат:', self.out, self.pick_out)
        ttk.Label(frm, foreground='#666',
                  text='Бренд каждого прайса определяется по кодам. Если ошибся — выделите прайс и нажмите '
                       '«Сменить бренд».').grid(row=4, column=0, columnspan=3, sticky='w', pady=(2, 0))

    def _label(self, p):
        name = BRANDS[self.brand.get(p, 'bohemia')][0]
        return f'{name}  —  {os.path.basename(p)}'

    def _added(self, ps):
        for p in ps:
            if p not in self.brand:
                try:
                    self.brand[p] = guzzini_photos.detect_brand(p)
                except Exception:
                    self.brand[p] = 'bohemia'
        self.prices.refresh()
        if not self.catalogs.get():
            self.catalogs.set(os.path.dirname(ps[0]))
        if not self.out.get():
            self.out.set(os.path.join(os.path.dirname(ps[0]), 'photos_result'))

    def toggle_brand(self):
        sel = self.prices.box.curselection()
        if not sel:
            messagebox.showinfo(APP_TITLE, 'Выделите прайс в списке, потом нажмите «Сменить бренд».')
            return
        for i in sel:
            p = self.prices.paths[i]
            self.brand[p] = 'guzzini' if self.brand.get(p) == 'bohemia' else 'bohemia'
        self.prices.refresh()

    def pick_catalogs(self):
        p = filedialog.askdirectory(title='Папка с PDF-каталогами')
        if p:
            self.catalogs.set(p)

    def pick_existing(self):
        p = filedialog.askdirectory(title='Папка с уже имеющимися фото')
        if p:
            self.existing.set(p)

    def pick_out(self):
        p = filedialog.askdirectory(title='Папка для результата')
        if p:
            self.out.set(p)

    def result_path(self):
        return self.out.get().strip()

    def job(self):
        prices, cats, out = list(self.prices.paths), self.catalogs.get().strip(), self.out.get().strip()
        if not prices:
            messagebox.showwarning(APP_TITLE, 'Нажмите «Добавить…» и выберите один или несколько прайс-листов.')
            return None
        if not os.path.isdir(cats):
            messagebox.showwarning(APP_TITLE, 'Выберите папку с PDF-каталогами.')
            return None
        if not out:
            messagebox.showwarning(APP_TITLE, 'Выберите папку для результата.')
            return None
        existing = self.existing.get().strip() or None
        if existing and not os.path.isdir(existing):
            messagebox.showwarning(APP_TITLE, 'Папка с имеющимися фото не найдена.')
            return None
        brands = {p: self.brand.get(p, 'bohemia') for p in prices}

        def work(log, progress):
            lines, last = [], out
            for n, p in enumerate(prices):
                name, module = BRANDS[brands[p]]
                stem = os.path.splitext(os.path.basename(p))[0]
                dest = out if len(prices) == 1 else os.path.join(out, f'{name}_{stem}')
                log(f'\n=== {name}: {os.path.basename(p)} ===')

                def prog(v, n=n):
                    progress((n + v) / len(prices))
                try:
                    stats, saved = module.run(p, cats, dest, log=log, progress=prog, existing_dir=existing)
                except core.PriceError as e:
                    log(f'Ошибка: {e}')
                    lines.append(f'{name} — {os.path.basename(p)}: ошибка: {e}')
                    continue
                found = sum(v for k, v in stats.items() if k not in ('есть в папке', 'нет в каталогах'))
                check = sum(v for k, v in stats.items() if k.startswith('проверить'))
                lines.append(f'{name} — {os.path.basename(p)}: артикулов {sum(stats.values())}, '
                             f'новых фото {found}, проверить {check}, уже есть {stats.get("есть в папке", 0)}, '
                             f'нет в каталогах {stats.get("нет в каталогах", 0)}')
                last = dest
            return 'Готово.\n\n' + '\n\n'.join(lines) + f'\n\nРезультат: {out}', out if len(prices) > 1 else last
        return work


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('860x640')
        self.minsize(660, 500)
        self.q = queue.Queue()
        self.worker = None
        self.last_out = None
        style = ttk.Style(self)
        style.configure('Big.TButton', font=('Segoe UI', 13, 'bold'), padding=(18, 14))
        style.configure('Run.TButton', font=('Segoe UI', 11, 'bold'), padding=(14, 6))
        style.configure('H.TLabel', font=('Segoe UI', 14, 'bold'))

        self.container = ttk.Frame(self)
        self.container.pack(fill='both', expand=True)
        self.home = self._build_home()
        self.screens = {'docs': DocsScreen, 'photos': PhotosScreen}
        self.work_frame = self._build_work()
        self.current = None
        self.show_home()
        self.after(100, self.poll)

    # ---- экраны
    def _build_home(self):
        f = ttk.Frame(self.container, padding=30)
        ttk.Label(f, text='Что нужно сделать?', style='H.TLabel').pack(pady=(10, 24))
        ttk.Button(f, text='📄  Proforma / Invoice  →  Excel', style='Big.TButton',
                   command=lambda: self.show('docs')).pack(fill='x', padx=60)
        ttk.Label(f, foreground='#555', justify='center',
                  text='Выписать позиции из PDF (или Excel) документов в Excel-отчёт\n'
                       'и, если нужно, сверить коды и цены с прайс-листами').pack(pady=(6, 26))
        ttk.Button(f, text='🖼  Фото из каталогов', style='Big.TButton',
                   command=lambda: self.show('photos')).pack(fill='x', padx=60)
        ttk.Label(f, foreground='#555', justify='center',
                  text='Вырезать фото товаров из PDF-каталогов по кодам из прайс-листов\n'
                       '(Bohemia, Guzzini; несколько прайсов за один запуск)').pack(pady=(6, 0))
        return f

    def _build_work(self):
        f = ttk.Frame(self.container, padding=(10, 8))
        top = ttk.Frame(f)
        top.pack(fill='x')
        self.back_btn = ttk.Button(top, text='← Назад', command=self.show_home)
        self.back_btn.pack(side='left')
        self.heading = ttk.Label(top, text='', style='H.TLabel')
        self.heading.pack(side='left', padx=14)
        self.body = ttk.Frame(f)
        self.body.pack(fill='x', pady=(10, 4))
        bar = ttk.Frame(f)
        bar.pack(fill='x', pady=6)
        self.btn = ttk.Button(bar, text='Запустить', style='Run.TButton', command=self.start)
        self.btn.pack(side='left')
        self.open_btn = ttk.Button(bar, text='Открыть результат', command=self.open_out, state='disabled')
        self.open_btn.pack(side='left', padx=8)
        self.status = ttk.Label(bar, text='')
        self.status.pack(side='right')
        self.pb = ttk.Progressbar(f, mode='determinate', maximum=100)
        self.pb.pack(fill='x')
        self.log = tk.Text(f, height=12, wrap='word', state='disabled')
        self.log.pack(fill='both', expand=True, pady=(8, 0))
        self.built = {}
        return f

    def show_home(self):
        if self.worker and self.worker.is_alive():
            return
        self.work_frame.pack_forget()
        self.home.pack(fill='both', expand=True)

    def show(self, key):
        self.home.pack_forget()
        if self.current is not None:
            self.current.pack_forget()
        if key not in self.built:
            self.built[key] = self.screens[key](self.body)
        self.current = self.built[key]
        self.current.pack(fill='x')
        self.heading.configure(text=self.current.heading)
        self.btn.configure(text=self.current.action)
        self.open_btn.configure(state='disabled')
        self.status.configure(text='')
        self.pb['value'] = 0
        self._clear_log()
        self.work_frame.pack(fill='both', expand=True)

    # ---- запуск
    def open_out(self):
        open_path(self.last_out)

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        work = self.current.job()
        if not work:
            return
        self.btn.configure(state='disabled')
        self.back_btn.configure(state='disabled')
        self.open_btn.configure(state='disabled')
        self.pb['value'] = 0
        self.status.configure(text='Работаю…')
        self._clear_log()
        self.worker = threading.Thread(target=self._work, args=(work,), daemon=True)
        self.worker.start()

    def _work(self, work):
        try:
            msg, out = work(lambda m: self.q.put(('log', m)), lambda v: self.q.put(('prog', v)))
            self.q.put(('done', (msg, out)))
        except (core.PriceError, invoice_reader.InvoiceError) as e:
            self.q.put(('log', f'\nОшибка: {e}'))
            self.q.put(('error', str(e)))
        except Exception as e:
            self.q.put(('log', traceback.format_exc()))
            self.q.put(('error', str(e)))

    def poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == 'log':
                    self._append(val)
                elif kind == 'prog':
                    self.pb['value'] = val * 100
                elif kind == 'done':
                    msg, self.last_out = val
                    self.pb['value'] = 100
                    self.status.configure(text='Готово')
                    self.btn.configure(state='normal')
                    self.back_btn.configure(state='normal')
                    self.open_btn.configure(state='normal')
                    messagebox.showinfo(APP_TITLE, msg)
                elif kind == 'error':
                    self.status.configure(text='Ошибка')
                    self.btn.configure(state='normal')
                    self.back_btn.configure(state='normal')
                    messagebox.showerror(APP_TITLE, f'Ошибка: {val}')
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def _append(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def _clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')


if __name__ == '__main__':
    App().mainloop()
