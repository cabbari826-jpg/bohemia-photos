#!/usr/bin/env python3
"""Окно для bohemia_photos: выбор прайса, папки каталогов, папки результата."""
import os
import sys
import queue
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import bohemia_photos as core

APP_TITLE = 'Bohemia — фото из каталогов'


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('760x560')
        self.minsize(600, 420)
        self.q = queue.Queue()
        self.worker = None

        self.price = tk.StringVar()
        self.catalogs = tk.StringVar()
        self.out = tk.StringVar()
        self.existing = tk.StringVar()

        pad = dict(padx=10, pady=6)
        frm = ttk.Frame(self)
        frm.pack(fill='x', **pad)
        frm.columnconfigure(1, weight=1)

        self._row(frm, 0, 'Прайс-лист (.xlsx, .xls):', self.price, self.pick_price)
        self._row(frm, 1, 'Папка с каталогами (PDF):', self.catalogs, self.pick_catalogs)
        self._row(frm, 2, 'Имеющиеся фото (необязательно):', self.existing, self.pick_existing)
        self._row(frm, 3, 'Куда сохранить результат:', self.out, self.pick_out)

        bar = ttk.Frame(self)
        bar.pack(fill='x', **pad)
        self.btn = ttk.Button(bar, text='Запустить', command=self.start)
        self.btn.pack(side='left')
        self.open_btn = ttk.Button(bar, text='Открыть папку результата', command=self.open_out, state='disabled')
        self.open_btn.pack(side='left', padx=8)
        self.status = ttk.Label(bar, text='')
        self.status.pack(side='right')

        self.pb = ttk.Progressbar(self, mode='determinate', maximum=100)
        self.pb.pack(fill='x', padx=10)

        self.log = tk.Text(self, height=16, wrap='word', state='disabled')
        self.log.pack(fill='both', expand=True, padx=10, pady=10)

        self.after(100, self.poll)

    def _row(self, parent, r, label, var, cmd):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky='w', pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky='ew', padx=6)
        ttk.Button(parent, text='Обзор…', command=cmd).grid(row=r, column=2)

    # ---- выбор путей
    def pick_price(self):
        p = filedialog.askopenfilename(title='Прайс-лист', filetypes=[('Excel', '*.xlsx *.xls *.csv'), ('Все файлы', '*.*')])
        if p:
            self.price.set(p)
            if not self.catalogs.get():
                self.catalogs.set(os.path.dirname(p))
            if not self.out.get():
                self.out.set(os.path.join(os.path.dirname(p), 'bohemia_result'))

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

    def open_out(self):
        p = self.out.get()
        if os.path.isdir(p):
            if sys.platform.startswith('win'):
                os.startfile(p)
            elif sys.platform == 'darwin':
                os.system(f'open "{p}"')
            else:
                os.system(f'xdg-open "{p}"')

    # ---- запуск
    def start(self):
        price, cats, out = self.price.get().strip(), self.catalogs.get().strip(), self.out.get().strip()
        if not os.path.isfile(price):
            messagebox.showwarning(APP_TITLE, 'Выберите файл прайс-листа.')
            return
        if not os.path.isdir(cats):
            messagebox.showwarning(APP_TITLE, 'Выберите папку с PDF-каталогами.')
            return
        if not out:
            messagebox.showwarning(APP_TITLE, 'Выберите папку для результата.')
            return
        existing = self.existing.get().strip() or None
        if existing and not os.path.isdir(existing):
            messagebox.showwarning(APP_TITLE, 'Папка с имеющимися фото не найдена.')
            return
        self.btn.configure(state='disabled')
        self.open_btn.configure(state='disabled')
        self.pb['value'] = 0
        self.status.configure(text='Работаю…')
        self._clear_log()
        self.worker = threading.Thread(target=self._work, args=(price, cats, out, existing), daemon=True)
        self.worker.start()

    def _work(self, price, cats, out, existing):
        try:
            stats, saved = core.run(price, cats, out,
                                    log=lambda m: self.q.put(('log', m)),
                                    progress=lambda v: self.q.put(('prog', v)),
                                    existing_dir=existing)
            self.q.put(('done', (stats, saved, out)))
        except core.PriceError as e:
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
                    stats, saved, out = val
                    self.pb['value'] = 100
                    self.status.configure(text='Готово')
                    self.btn.configure(state='normal')
                    self.open_btn.configure(state='normal')
                    found = stats.get('совпадает', 0) + stats.get('проверить', 0)
                    messagebox.showinfo(APP_TITLE,
                                        f'Готово. Артикулов: {sum(stats.values())}\n\n'
                                        f'Уже есть в папке: {stats.get("есть в папке", 0)}\n'
                                        f'Новых фото из каталогов: {found}\n'
                                        f'Проверить вручную: {stats.get("проверить", 0)}\n'
                                        f'Нет в каталогах: {stats.get("нет в каталогах", 0)}\n\n'
                                        f'Отчёт: {os.path.join(out, "report.xlsx")}')
                elif kind == 'error':
                    self.status.configure(text='Ошибка')
                    self.btn.configure(state='normal')
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
