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
import updater

APP_TITLE = 'Inventa — каталоги и документы'
BRANDS = {'bohemia': ('Bohemia', core), 'guzzini': ('Guzzini', guzzini_photos), 'helios': ('Helios', guzzini_photos)}
PRICE_TYPES = [('Excel и PDF', '*.xlsx *.xls *.csv *.pdf *.PDF'), ('Все файлы', '*.*')]


def open_path(p):
    if not p or not os.path.exists(p):
        return
    if sys.platform.startswith('win'):
        os.startfile(p)
    elif sys.platform == 'darwin':
        os.system(f'open "{p}"')
    else:
        os.system(f'xdg-open "{p}"')


# ---- перетаскивание файлов мышью (tkinterdnd2; без него программа работает как раньше)
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:          # библиотеки нет — просто без drag'n'drop
    TkinterDnD = DND_FILES = None
DND_OK = False


def enable_dnd(root):
    global DND_OK
    if TkinterDnD is None:
        return
    try:
        TkinterDnD._require(root)
        DND_OK = True
    except Exception:
        DND_OK = False


def on_drop(widget, callback):
    """callback(список путей) при перетаскивании файлов/папок на виджет."""
    if not DND_OK:
        return

    def handler(e):
        try:
            paths = [p for p in widget.tk.splitlist(e.data) if p]
            if paths:
                callback(paths)
        except Exception:
            traceback.print_exc()
        return e.action

    try:
        widget.drop_target_register(DND_FILES)
        widget.dnd_bind('<<Drop>>', handler)
    except Exception:
        pass


def path_row(parent, r, label, var, cmd, folder=True):
    ttk.Label(parent, text=label).grid(row=r, column=0, sticky='w', pady=4)
    e = ttk.Entry(parent, textvariable=var)
    e.grid(row=r, column=1, sticky='ew', padx=6)
    ttk.Button(parent, text='Обзор…', command=cmd).grid(row=r, column=2, sticky='ew')
    # перетащили файл в поле папки — берём папку, где он лежит
    on_drop(e, lambda ps: var.set(os.path.normpath(ps[0] if not folder or os.path.isdir(ps[0])
                                                    else os.path.dirname(ps[0]))))
    return e


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
        on_drop(self.box, self.add_paths)
        if DND_OK and not self.paths:
            self.box.insert('end', '  (перетащите файлы сюда или нажмите «Добавить…»)')
            self.box.itemconfigure(0, foreground='#999')
            self._hint = True

    _hint = False

    def _exts(self):
        ext = set()
        for _, pat in self.filetypes:
            for x in pat.split():
                if x in ('*.*', '*'):
                    continue
                ext.add(x.lstrip('*').lower())
            break      # первая группа — «свои» типы файлов
        return ext

    def add_paths(self, ps):
        """Перетащенные файлы и папки (из папки — файлы подходящих типов)."""
        ext = self._exts()
        files = []
        for p in ps:
            p = os.path.normpath(p)
            if os.path.isdir(p):
                files += sorted(os.path.join(p, f) for f in os.listdir(p)
                                if os.path.splitext(f)[1].lower() in ext and not f.startswith('~$'))
            elif os.path.isfile(p):
                files.append(p)
        skipped = [f for f in files if ext and os.path.splitext(f)[1].lower() not in ext]
        files = [f for f in files if f not in skipped]
        self._add(files)
        if skipped:
            messagebox.showwarning(APP_TITLE, 'Не добавлены (другой тип файла):\n' +
                                   '\n'.join(os.path.basename(f) for f in skipped[:15]))

    def add(self):
        self._add(filedialog.askopenfilenames(title=self.title, filetypes=self.filetypes))

    def _add(self, ps):
        if ps and self._hint:
            self.box.delete(0, 'end')
            self._hint = False
        for p in ps:
            if p not in self.paths:
                self.paths.append(p)
                self.box.insert('end', self.label(p))
        if ps and self.on_add:
            self.on_add(ps)

    def remove(self):
        if self._hint:
            return
        for i in reversed(self.box.curselection()):
            self.box.delete(i)
            del self.paths[i]

    def refresh(self):
        if self._hint:
            return
        sel = self.box.curselection()
        self.box.delete(0, 'end')
        for p in self.paths:
            self.box.insert('end', self.label(p))
        for i in sel:
            self.box.selection_set(i)


class DocsScreen(ttk.Frame):
    """Proforma / Invoice -> Excel."""
    action = 'Сделать Excel'
    heading = 'Proforma / Invoice → Excel, сверка с заказом'

    def __init__(self, master):
        super().__init__(master)
        self.out = tk.StringVar()
        frm = ttk.Frame(self)
        frm.pack(fill='x')
        frm.columnconfigure(1, weight=1)
        ttk.Label(frm, text='1. Документы поставщика\n(Proforma / Invoice, PDF или Excel):').grid(
            row=0, column=0, sticky='nw', pady=4)
        self.docs = FileList(frm, 'Proforma / Invoice',
                             [('PDF и Excel', '*.pdf *.PDF *.xlsx *.xls *.csv'), ('Все файлы', '*.*')], height=3)
        self.docs.grid(row=0, column=1, columnspan=2, sticky='ew', padx=6)
        self.docs.on_add = self._suggest_out
        ttk.Label(frm, text='2. Наши заказы поставщику\n(Excel, необязательно):').grid(
            row=1, column=0, sticky='nw', pady=4)
        self.orders = FileList(frm, 'Заказы поставщику', [('Excel', '*.xlsx *.xls *.csv'), ('Все файлы', '*.*')],
                               height=3)
        self.orders.grid(row=1, column=1, columnspan=2, sticky='ew', padx=6, pady=4)
        self.orders.on_add = self._suggest_out
        ttk.Label(frm, text='3. Прайсы для сверки\n(необязательно, можно несколько):').grid(
            row=2, column=0, sticky='nw', pady=4)
        self.prices = FileList(frm, 'Прайс-листы', PRICE_TYPES, height=3)
        self.prices.grid(row=2, column=1, columnspan=2, sticky='ew', padx=6, pady=4)
        ttk.Label(frm, text='4. Куда сохранить Excel:').grid(row=3, column=0, sticky='w', pady=4)
        out_e = ttk.Entry(frm, textvariable=self.out)
        out_e.grid(row=3, column=1, sticky='ew', padx=6)
        ttk.Button(frm, text='Обзор…', command=self.pick_out).grid(row=3, column=2, sticky='ew')
        # перетащили папку — отчёт в неё, перетащили xlsx — в этот файл
        on_drop(out_e, lambda ps: self.out.set(os.path.join(ps[0], 'documents_report.xlsx') if os.path.isdir(ps[0])
                                               else os.path.normpath(ps[0])))
        ttk.Label(frm, foreground='#666',
                  text='Если добавить и документы, и заказы — в отчёте будет лист «Сверка с заказом»: '
                       'чего нет, что лишнее, где другое кол-во или цена.', wraplength=760, justify='left').grid(
            row=4, column=0, columnspan=3, sticky='w', pady=(2, 0))

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
        orders = list(self.orders.paths)
        if not files and not orders:
            messagebox.showwarning(APP_TITLE, 'Добавьте документы поставщика (Proforma / Invoice) '
                                              'и/или наши заказы — кнопки «Добавить…».')
            return None
        if not out:
            messagebox.showwarning(APP_TITLE, 'Укажите, куда сохранить Excel.')
            return None
        if not out.lower().endswith('.xlsx'):
            out += '.xlsx'
            self.out.set(out)

        def work(log, progress):
            r = invoice_reader.run(files, out, prices or None, log=log, progress=progress, order_paths=orders or None)
            msg = f'Готово. Документов: {r["docs"]}, позиций: {r["items"]}\n'
            if r['orders']:
                msg += f'Позиций в заказах: {r["orders"]}\n'
            if r['order_stats']:
                msg += '\nСверка заказа с документами:\n' + '\n'.join(
                    f'  {k}: {v}' for k, v in r['order_stats'].most_common()) + '\n'
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
        sel = [] if self.prices._hint else self.prices.box.curselection()
        if not sel:
            messagebox.showinfo(APP_TITLE, 'Выделите прайс в списке, потом нажмите «Сменить бренд».')
            return
        for i in sel:
            p = self.prices.paths[i]
            order = list(BRANDS)
            self.brand[p] = order[(order.index(self.brand.get(p, 'bohemia')) + 1) % len(order)]
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


def enable_dpi_awareness():
    """Windows с масштабом 125–175 %: без этого окно растягивается системой, выходит за экран и размыто."""
    if sys.platform.startswith('win'):
        try:
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


class App(tk.Tk):
    def __init__(self):
        enable_dpi_awareness()
        super().__init__()
        enable_dnd(self)
        self.title(APP_TITLE)
        try:   # иконка окна (в exe файл вшит рядом с программой)
            base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
            self.iconbitmap(default=os.path.join(base, 'app.ico'))
        except Exception:
            pass
        self.place_window()
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
        if getattr(sys, 'frozen', False):
            self.after(300, self.after_update)
            self.after(1500, self.check_updates)    # тихая проверка при запуске собранной программы

    def place_window(self):
        """Окно по центру экрана и целиком на экране; на маленьком экране — развернуть."""
        self.update_idletasks()
        scale = max(1.0, self.winfo_fpixels('1i') / 96)        # масштаб Windows (1.25, 1.5…)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = min(int(900 * scale), int(sw * 0.9))
        h = min(int(680 * scale), int(sh * 0.85))               # запас на панель задач
        self.minsize(min(int(640 * scale), w), min(int(480 * scale), h))
        x, y = max(0, (sw - w) // 2), max(0, (sh - h) // 2 - int(20 * scale))
        self.geometry(f'{w}x{h}+{x}+{y}')
        if sh < 800 * scale and sys.platform.startswith('win'):
            try:
                self.state('zoomed')
            except tk.TclError:
                pass

    # ---- экраны
    def _build_home(self):
        f = ttk.Frame(self.container, padding=30)
        ttk.Label(f, text='Что нужно сделать?', style='H.TLabel').pack(pady=(10, 24))
        ttk.Button(f, text='📄  Proforma / Invoice  →  Excel', style='Big.TButton',
                   command=lambda: self.show('docs')).pack(fill='x', padx=60)
        ttk.Label(f, foreground='#555', justify='center',
                  text='Выписать позиции из PDF (или Excel) документов в Excel-отчёт,\n'
                       'сверить их с нашим заказом поставщику и с прайс-листами').pack(pady=(6, 26))
        ttk.Button(f, text='🖼  Фото из каталогов', style='Big.TButton',
                   command=lambda: self.show('photos')).pack(fill='x', padx=60)
        ttk.Label(f, foreground='#555', justify='center',
                  text='Вырезать фото товаров из PDF-каталогов по кодам из прайс-листов\n'
                       '(Bohemia, Guzzini, Helios; несколько прайсов за один запуск)').pack(pady=(6, 0))
        bottom = ttk.Frame(f)
        bottom.pack(side='bottom', fill='x')
        ttk.Label(bottom, text=f'Версия {updater.current_version()}  ·  © Jooble Soft', foreground='#777').pack(side='left')
        self.upd_btn = ttk.Button(bottom, text='Проверить обновления', command=lambda: self.check_updates(True))
        self.upd_btn.pack(side='right')
        self.upd_label = ttk.Label(bottom, text='', foreground='#777')
        self.upd_label.pack(side='right', padx=8)
        return f

    # ---- обновления
    def after_update(self):
        """Проверка после автообновления: версия сменилась или нет (тогда — журнал установщика)."""
        try:
            res = updater.pending_result()
        except Exception:
            return
        if not res:
            return
        if res[0] == 'ok':
            self.upd_label.configure(text=f'Обновлено до версии {res[1]}')
            return
        if messagebox.askyesno(APP_TITLE, f'Обновление до версии {res[1]} не установилось '
                                          f'(сейчас {updater.current_version()}).\n\n'
                                          f'Конец журнала установки:\n{res[2][-1200:]}\n\n'
                                          f'Открыть страницу загрузки, чтобы поставить вручную?'):
            import webbrowser
            webbrowser.open(f'https://github.com/{updater.REPO}/releases/latest')

    def check_updates(self, manual=False):
        self.upd_btn.configure(state='disabled')
        self.upd_label.configure(text='Проверяю…' if manual else '')

        def worker():
            try:
                self.q.put(('upd', (updater.check(), manual, None)))
            except Exception as e:
                self.q.put(('upd', (None, manual, e)))
        threading.Thread(target=worker, daemon=True).start()

    def _on_update(self, rel, manual, err):
        self.upd_btn.configure(state='normal')
        if err:
            self.upd_label.configure(text='')
            if manual:
                if getattr(err, 'code', None) == 404:
                    messagebox.showwarning(APP_TITLE, 'Не удалось найти релизы: репозиторий программы на GitHub '
                                                      'закрыт (private), а программа заходит туда без входа. '
                                                      'Сделайте репозиторий публичным, и обновления заработают.')
                else:
                    messagebox.showwarning(APP_TITLE, f'Не удалось проверить обновления (нет интернета или GitHub '
                                                      f'недоступен).\n\n{err}')
            return
        if not rel:
            self.upd_label.configure(text='Установлена последняя версия' if manual else '')
            return
        self.upd_label.configure(text=f'Есть версия {rel["version"]}')
        notes = rel['notes'].strip()
        if len(notes) > 900:
            notes = notes[:900] + '…'
        if updater.is_installed() and rel['setup_url']:
            if messagebox.askyesno(APP_TITLE, f'Доступна новая версия {rel["version"]} '
                                              f'(у вас {updater.current_version()}).\n\n{notes}\n\n'
                                              f'Скачать и установить сейчас? Программа закроется и откроется снова.'):
                self.install_update(rel)
        elif messagebox.askyesno(APP_TITLE, f'Доступна новая версия {rel["version"]} '
                                            f'(у вас {updater.current_version()}).\n\n'
                                            f'Эта копия запущена без установки. Открыть страницу загрузки?'):
            import webbrowser
            webbrowser.open(rel['page'])

    def install_update(self, rel):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo(APP_TITLE, 'Дождитесь окончания текущей работы и повторите.')
            return
        self.upd_btn.configure(state='disabled')
        self.upd_label.configure(text='Скачиваю обновление…')
        self._upd_version = rel['version']

        def worker():
            try:
                path = updater.download(rel['setup_url'],
                                        lambda v: self.q.put(('upd_prog', v)))
                self.q.put(('upd_ready', path))
            except Exception as e:
                self.q.put(('upd_fail', str(e)))
        threading.Thread(target=worker, daemon=True).start()

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
                elif kind == 'upd':
                    self._on_update(*val)
                elif kind == 'upd_prog':
                    self.upd_label.configure(text=f'Скачиваю обновление… {int(val * 100)}%')
                elif kind == 'upd_ready':
                    self.upd_label.configure(text='Устанавливаю…')
                    try:
                        updater.run_installer(val, self._upd_version)
                    except Exception as e:
                        self.upd_btn.configure(state='normal')
                        self.upd_label.configure(text='')
                        messagebox.showerror(APP_TITLE, f'Не удалось запустить установщик: {e}\n\n'
                                                        f'Файл установщика: {val}\nЕго можно запустить вручную.')
                    else:
                        # выходим совсем, чтобы освободить Inventa.exe для замены
                        self.after(300, lambda: (self.destroy(), os._exit(0)))
                elif kind == 'upd_fail':
                    self.upd_btn.configure(state='normal')
                    self.upd_label.configure(text='')
                    messagebox.showerror(APP_TITLE, f'Не удалось скачать обновление: {val}')
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
