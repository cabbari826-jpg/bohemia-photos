#!/usr/bin/env python3
"""Окно программы: три вкладки.
  1. Bohemia — фото из каталогов по прайсу.
  2. Guzzini — фото из каталогов и журналов по прайсу.
  3. Proforma / Invoice — PDF или Excel документа -> Excel-отчёт и сверка с прайсом.
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


def open_path(p):
    if not p or not os.path.exists(p):
        return
    if sys.platform.startswith('win'):
        os.startfile(p)
    elif sys.platform == 'darwin':
        os.system(f'open "{p}"')
    else:
        os.system(f'xdg-open "{p}"')


class PhotosTab(ttk.Frame):
    """Вкладка «фото из каталогов» (Bohemia или Guzzini — отличаются только модулем)."""

    def __init__(self, master, app, module, default_out):
        super().__init__(master)
        self.app, self.module, self.default_out = app, module, default_out
        self.price, self.catalogs = tk.StringVar(), tk.StringVar()
        self.out, self.existing = tk.StringVar(), tk.StringVar()
        frm = ttk.Frame(self)
        frm.pack(fill='x', padx=10, pady=8)
        frm.columnconfigure(1, weight=1)
        app.row(frm, 0, 'Прайс-лист (.xlsx, .xls):', self.price, self.pick_price)
        app.row(frm, 1, 'Папка с каталогами (PDF):', self.catalogs, self.pick_catalogs)
        app.row(frm, 2, 'Имеющиеся фото (необязательно):', self.existing, self.pick_existing)
        app.row(frm, 3, 'Куда сохранить результат:', self.out, self.pick_out)

    def pick_price(self):
        p = filedialog.askopenfilename(title='Прайс-лист',
                                       filetypes=[('Excel', '*.xlsx *.xls *.csv'), ('Все файлы', '*.*')])
        if p:
            self.price.set(p)
            if not self.catalogs.get():
                self.catalogs.set(os.path.dirname(p))
            if not self.out.get():
                self.out.set(os.path.join(os.path.dirname(p), self.default_out))

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
        """Проверяет поля и возвращает функцию работы (или None)."""
        price, cats, out = self.price.get().strip(), self.catalogs.get().strip(), self.out.get().strip()
        if not os.path.isfile(price):
            messagebox.showwarning(APP_TITLE, 'Выберите файл прайс-листа.')
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

        def work(log, progress):
            stats, saved = self.module.run(price, cats, out, log=log, progress=progress, existing_dir=existing)
            found = sum(v for k, v in stats.items() if k not in ('есть в папке', 'нет в каталогах'))
            check = sum(v for k, v in stats.items() if k.startswith('проверить'))
            return (f'Готово. Артикулов: {sum(stats.values())}\n\n'
                    f'Уже есть в папке: {stats.get("есть в папке", 0)}\n'
                    f'Новых фото из каталогов: {found}\n'
                    f'Проверить вручную: {check}\n'
                    f'Нет в каталогах: {stats.get("нет в каталогах", 0)}\n\n'
                    f'Отчёт: {os.path.join(out, "report.xlsx")}')
        return work


class InvoiceTab(ttk.Frame):
    """Вкладка Proforma / Invoice: документы -> Excel, сверка с прайсом."""

    def __init__(self, master, app):
        super().__init__(master)
        self.app = app
        self.files = []
        self.price, self.out = tk.StringVar(), tk.StringVar()
        frm = ttk.Frame(self)
        frm.pack(fill='x', padx=10, pady=8)
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text='Документы (PDF, Excel):').grid(row=0, column=0, sticky='nw', pady=4)
        box = ttk.Frame(frm)
        box.grid(row=0, column=1, sticky='ew', padx=6)
        box.columnconfigure(0, weight=1)
        self.listbox = tk.Listbox(box, height=4, selectmode='extended')
        self.listbox.grid(row=0, column=0, sticky='ew')
        btns = ttk.Frame(frm)
        btns.grid(row=0, column=2, sticky='n')
        ttk.Button(btns, text='Добавить…', command=self.add_files).pack(fill='x')
        ttk.Button(btns, text='Убрать', command=self.remove_files).pack(fill='x', pady=2)
        app.row(frm, 1, 'Прайс для сверки (необязательно):', self.price, self.pick_price)
        app.row(frm, 2, 'Куда сохранить отчёт (.xlsx):', self.out, self.pick_out)

    def add_files(self):
        ps = filedialog.askopenfilenames(title='Proforma / Invoice',
                                         filetypes=[('PDF и Excel', '*.pdf *.PDF *.xlsx *.xls *.csv'),
                                                    ('Все файлы', '*.*')])
        for p in ps:
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert('end', os.path.basename(p))
        if ps and not self.out.get():
            self.out.set(os.path.join(os.path.dirname(ps[0]), 'documents_report.xlsx'))

    def remove_files(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]

    def pick_price(self):
        p = filedialog.askopenfilename(title='Прайс-лист',
                                       filetypes=[('Excel', '*.xlsx *.xls *.csv'), ('Все файлы', '*.*')])
        if p:
            self.price.set(p)

    def pick_out(self):
        p = filedialog.asksaveasfilename(title='Отчёт', defaultextension='.xlsx',
                                         filetypes=[('Excel', '*.xlsx')], initialfile='documents_report.xlsx')
        if p:
            self.out.set(p)

    def result_path(self):
        return os.path.dirname(self.out.get().strip())

    def job(self):
        files, price, out = list(self.files), self.price.get().strip() or None, self.out.get().strip()
        if not files:
            messagebox.showwarning(APP_TITLE, 'Добавьте PDF или Excel файлы Proforma / Invoice.')
            return None
        if price and not os.path.isfile(price):
            messagebox.showwarning(APP_TITLE, 'Файл прайса не найден.')
            return None
        if not out:
            messagebox.showwarning(APP_TITLE, 'Укажите, куда сохранить отчёт.')
            return None
        if not out.lower().endswith('.xlsx'):
            out += '.xlsx'
            self.out.set(out)

        def work(log, progress):
            r = invoice_reader.run(files, out, price, log=log, progress=progress)
            msg = f'Готово. Документов: {r["docs"]}, позиций: {r["items"]}\n'
            if r['bad_lines']:
                msg += f'Строк, где кол-во × цена ≠ сумма: {r["bad_lines"]}\n'
            if r['stats']:
                msg += '\nСверка с прайсом:\n' + '\n'.join(f'  {k}: {v}' for k, v in r['stats'].most_common())
            if r['errors']:
                msg += '\n\nНе прочитаны:\n' + '\n'.join(r['errors'])
            return msg + f'\n\nОтчёт: {out}'
        return work


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('820x620')
        self.minsize(640, 480)
        self.q = queue.Queue()
        self.worker = None

        self.tabs = ttk.Notebook(self)
        self.tabs.pack(fill='x', padx=10, pady=(10, 0))
        self.tab_bohemia = PhotosTab(self.tabs, self, core, 'bohemia_result')
        self.tab_guzzini = PhotosTab(self.tabs, self, guzzini_photos, 'guzzini_result')
        self.tab_docs = InvoiceTab(self.tabs, self)
        self.tabs.add(self.tab_bohemia, text='Bohemia — фото')
        self.tabs.add(self.tab_guzzini, text='Guzzini — фото')
        self.tabs.add(self.tab_docs, text='Proforma / Invoice')

        bar = ttk.Frame(self)
        bar.pack(fill='x', padx=10, pady=6)
        self.btn = ttk.Button(bar, text='Запустить', command=self.start)
        self.btn.pack(side='left')
        self.open_btn = ttk.Button(bar, text='Открыть папку результата', command=self.open_out, state='disabled')
        self.open_btn.pack(side='left', padx=8)
        self.status = ttk.Label(bar, text='')
        self.status.pack(side='right')

        self.pb = ttk.Progressbar(self, mode='determinate', maximum=100)
        self.pb.pack(fill='x', padx=10)

        self.log = tk.Text(self, height=14, wrap='word', state='disabled')
        self.log.pack(fill='both', expand=True, padx=10, pady=10)
        self.last_out = None
        self.after(100, self.poll)

    @staticmethod
    def row(parent, r, label, var, cmd):
        ttk.Label(parent, text=label).grid(row=r, column=0, sticky='w', pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=r, column=1, sticky='ew', padx=6)
        ttk.Button(parent, text='Обзор…', command=cmd).grid(row=r, column=2)

    def current(self):
        return self.tabs.nametowidget(self.tabs.select())

    def open_out(self):
        open_path(self.last_out)

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        tab = self.current()
        work = tab.job()
        if not work:
            return
        self.last_out = tab.result_path()
        self.btn.configure(state='disabled')
        self.open_btn.configure(state='disabled')
        self.pb['value'] = 0
        self.status.configure(text='Работаю…')
        self._clear_log()
        self.worker = threading.Thread(target=self._work, args=(work,), daemon=True)
        self.worker.start()

    def _work(self, work):
        try:
            msg = work(lambda m: self.q.put(('log', m)), lambda v: self.q.put(('prog', v)))
            self.q.put(('done', msg))
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
                    self.pb['value'] = 100
                    self.status.configure(text='Готово')
                    self.btn.configure(state='normal')
                    self.open_btn.configure(state='normal')
                    messagebox.showinfo(APP_TITLE, val)
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
