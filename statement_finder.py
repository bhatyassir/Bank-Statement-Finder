"""Statement Finder: finds bank statements hidden in a folder and copies them, bank by bank, into another folder.

    Double-click "Statement Finder.bat", or run:   python statement_finder.py

Your originals are never moved, renamed or changed: files are only copied.
The engine is in finder_core.py. This file is only the window.
"""
import json
import multiprocessing as mp
import os
import queue
import sys
import threading
import time
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import finder_core as core

APP_NAME = 'Statement Finder'
SETTINGS_DIR = os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), 'StatementFinder')
SETTINGS_FILE = os.path.join(SETTINGS_DIR, 'settings.json')
CRASH_FILE = os.path.join(SETTINGS_DIR, 'crash.log')


def nice_time(seconds):
    seconds = int(seconds)
    if seconds < 60:
        return f'{seconds} second{"s" if seconds != 1 else ""}'
    m, s = divmod(seconds, 60)
    if m < 60:
        return f'{m} min {s:02d} s'
    h, m = divmod(m, 60)
    return f'{h} h {m:02d} min'


class App:
    def __init__(self, root):
        self.root = root
        root.title(APP_NAME)
        self.events = queue.Queue()
        self.stop = threading.Event()
        self.thread = None
        self.running = False
        self.result = None
        self.counts = {}
        self.sure = self.not_sure = self.copied = 0
        self.last_report = ''
        self.active_source = ''
        self.active_dry = False
        self._style()
        self._build()
        self._load_settings()
        self.root.protocol('WM_DELETE_WINDOW', self._close)
        # the window is as big as its content needs (text is larger on a scaled screen), but never bigger than the screen
        root.update_idletasks()
        scale = root.winfo_fpixels('1i') / 96.0
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        w = min(max(root.winfo_reqwidth(), int(940 * scale)), sw - 20)
        h = min(root.winfo_reqheight() + 8, sh - int(90 * scale))
        root.minsize(min(w, int(820 * scale)), min(h, int(640 * scale)))
        root.geometry(f'{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 4)}')
        self._poll()

    # ------------------------------------------------------------------------------------------------- look
    def _style(self):
        st = ttk.Style()
        if 'vista' in st.theme_names():
            st.theme_use('vista')
        elif 'clam' in st.theme_names():
            st.theme_use('clam')
        self.font = ('Segoe UI', 10)
        st.configure('.', font=self.font)
        st.configure('Title.TLabel', font=('Segoe UI', 18, 'bold'))
        st.configure('Sub.TLabel', foreground='#555555')
        st.configure('Hint.TLabel', foreground='#666666', font=('Segoe UI', 9))
        st.configure('Go.TButton', font=('Segoe UI', 11, 'bold'), padding=(18, 7))
        st.configure('TButton', padding=(10, 5))
        st.configure('Treeview', rowheight=24)
        st.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'))

    def _build(self):
        pad = {'padx': 14}
        outer = ttk.Frame(self.root)
        outer.pack(fill='both', expand=True)
        ttk.Label(outer, text='Statement Finder', style='Title.TLabel').pack(anchor='w', padx=14, pady=(12, 0))
        ttk.Label(outer, text='Looks through a folder and everything under it for bank statements, and copies them into one '
                              'folder per bank.', style='Sub.TLabel').pack(anchor='w', padx=14, pady=(0, 8))

        # 1. folders
        box = ttk.LabelFrame(outer, text=' 1. Folders ')
        box.pack(fill='x', **pad)
        box.columnconfigure(1, weight=1)
        self.source = tk.StringVar()
        self.dest = tk.StringVar()
        for r, (label, var, cmd) in enumerate((('Look in:', self.source, self._browse_source), ('Copy to:', self.dest, self._browse_dest))):
            ttk.Label(box, text=label, width=9).grid(row=r, column=0, sticky='w', padx=(10, 4), pady=6)
            ttk.Entry(box, textvariable=var).grid(row=r, column=1, sticky='ew', pady=6)
            ttk.Button(box, text='Browse...', command=cmd).grid(row=r, column=2, padx=10, pady=6)
        ttk.Label(box, text='Every sub-folder is searched. Your originals are never moved or changed: files are only copied.',
                  style='Hint.TLabel').grid(row=2, column=0, columnspan=3, sticky='w', padx=10, pady=(0, 8))

        # 2. options
        opt = ttk.LabelFrame(outer, text=' 2. What to look for ')
        opt.pack(fill='x', padx=14, pady=(8, 0))
        self.kind_vars = {}
        row = ttk.Frame(opt)
        row.pack(fill='x', padx=10, pady=(8, 2))
        ttk.Label(row, text='File types:').pack(side='left', padx=(0, 10))
        for key in ('pdf', 'word', 'text'):
            v = tk.BooleanVar(value=True)
            self.kind_vars[key] = v
            ttk.Checkbutton(row, text=core.FILE_TYPE_NAMES[key], variable=v).pack(side='left', padx=(0, 14))
        self.not_sure_var = tk.BooleanVar(value=True)
        self.system_var = tk.BooleanVar(value=True)
        self.dry_var = tk.BooleanVar(value=False)
        for var, text in ((self.not_sure_var, "Also copy the ones I am not sure about (they go in a \"(not sure)\" folder inside the bank's folder)"),
                          (self.system_var, 'Skip Windows, program and cache folders'),
                          (self.dry_var, 'Scan only: show what would be copied, but copy nothing')):
            ttk.Checkbutton(opt, text=text, variable=var).pack(anchor='w', padx=10, pady=1)
        row = ttk.Frame(opt)
        row.pack(fill='x', padx=10, pady=(6, 8))
        cpu = os.cpu_count() or 4
        self.workers = tk.IntVar(value=cpu)
        self.timeout = tk.IntVar(value=60)
        ttk.Label(row, text='Read').pack(side='left')
        ttk.Spinbox(row, from_=1, to=64, width=4, textvariable=self.workers).pack(side='left', padx=6)
        ttk.Label(row, text=f'files at the same time (this PC has {cpu} processor threads)').pack(side='left')
        ttk.Label(row, text='      Give up on a file after').pack(side='left')
        ttk.Spinbox(row, from_=10, to=600, increment=10, width=5, textvariable=self.timeout).pack(side='left', padx=6)
        ttk.Label(row, text='seconds').pack(side='left')

        # 3. go
        go = ttk.Frame(outer)
        go.pack(fill='x', padx=14, pady=10)
        self.start_btn = ttk.Button(go, text='Find statements', style='Go.TButton', command=self._start)
        self.start_btn.pack(side='left')
        self.stop_btn = ttk.Button(go, text='Stop', command=self._stop, state='disabled')
        self.stop_btn.pack(side='left', padx=8)
        self.open_btn = ttk.Button(go, text='Open the copies folder', command=self._open_dest, state='disabled')
        self.open_btn.pack(side='right')
        self.report_btn = ttk.Button(go, text='Open the report', command=self._open_report, state='disabled')
        self.report_btn.pack(side='right', padx=8)

        self.bar = ttk.Progressbar(outer, mode='determinate')
        self.bar.pack(fill='x', padx=14)
        self.status = tk.StringVar(value='Choose the two folders and press "Find statements".')
        ttk.Label(outer, textvariable=self.status).pack(anchor='w', padx=14, pady=(6, 0))
        self.counter = tk.StringVar(value='')
        ttk.Label(outer, textvariable=self.counter, style='Hint.TLabel').pack(anchor='w', padx=14)

        # results by bank
        res = ttk.LabelFrame(outer, text=' Found, by bank ')
        res.pack(fill='x', padx=14, pady=(8, 0))
        cols = ('bank', 'sure', 'maybe')
        self.tree = ttk.Treeview(res, columns=cols, show='headings', height=5)
        self.tree.heading('bank', text='Bank')
        self.tree.heading('sure', text='Statements')
        self.tree.heading('maybe', text='Not sure')
        self.tree.column('bank', width=420, anchor='w')
        self.tree.column('sure', width=110, anchor='e')
        self.tree.column('maybe', width=110, anchor='e')
        sb = ttk.Scrollbar(res, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='x', expand=True, padx=(8, 0), pady=8)
        sb.pack(side='left', fill='y', pady=8, padx=(0, 8))

        # log
        logf = ttk.LabelFrame(outer, text=' What was found ')
        logf.pack(fill='both', expand=True, padx=14, pady=(8, 0))
        self.log = tk.Text(logf, height=8, wrap='none', state='disabled', font=('Consolas', 9), background='#fbfbfb', relief='flat')
        ys = ttk.Scrollbar(logf, orient='vertical', command=self.log.yview)
        xs = ttk.Scrollbar(logf, orient='horizontal', command=self.log.xview)
        self.log.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        logf.rowconfigure(0, weight=1)
        logf.columnconfigure(0, weight=1)
        self.log.grid(row=0, column=0, sticky='nsew', padx=(8, 0), pady=(8, 0))
        ys.grid(row=0, column=1, sticky='ns', pady=(8, 0), padx=(0, 8))
        xs.grid(row=1, column=0, sticky='ew', padx=(8, 0), pady=(0, 8))
        self.log.tag_configure('sure', foreground='#0b5d1e')
        self.log.tag_configure('maybe', foreground='#9a5b00')
        self.log.tag_configure('locked', foreground='#555555')
        self.log.tag_configure('error', foreground='#b00020')

        readers = core.available_readers()
        bits = [f'{k}: {v}' if v else f'{k}: NOT INSTALLED' for k, v in readers.items()]
        missing = [k for k, v in readers.items() if not v]
        ttk.Label(outer, text='Readers  ' + '   '.join(bits), style='Hint.TLabel',
                  foreground='#b00020' if missing else '#666666').pack(anchor='w', padx=14, pady=(4, 8))

    # ------------------------------------------------------------------------------------------------- settings
    def _load_settings(self):
        try:
            with open(SETTINGS_FILE, encoding='utf-8') as f:
                s = json.load(f)
        except (OSError, ValueError):
            return
        self.source.set(s.get('source', ''))
        self.dest.set(s.get('dest', ''))
        saved_kinds = [k for k in s.get('kinds', list(self.kind_vars)) if k in self.kind_vars]
        if not saved_kinds:
            saved_kinds = list(self.kind_vars)
        for k, v in self.kind_vars.items():
            v.set(k in saved_kinds)
        self.not_sure_var.set(s.get('include_not_sure', True))
        self.system_var.set(s.get('skip_system', True))
        if s.get('timeout'):
            self.timeout.set(int(s['timeout']))

    @staticmethod
    def _safe_int(var, default, lo, hi=None):
        """Reads an IntVar Spinbox value, falling back to `default` if cleared or invalid."""
        try:
            n = int(var.get())
        except (tk.TclError, ValueError, TypeError):
            n = default
        return max(lo, min(hi, n) if hi is not None else n)

    def _save_settings(self):
        try:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump({'source': self.source.get(), 'dest': self.dest.get(),
                           'kinds': [k for k, v in self.kind_vars.items() if v.get()],
                           'include_not_sure': self.not_sure_var.get(),
                           'skip_system': self.system_var.get(),
                           'timeout': self._safe_int(self.timeout, 60, 10)}, f, indent=1)
        except OSError:
            pass

    # ------------------------------------------------------------------------------------------------- buttons
    def _browse_source(self):
        d = filedialog.askdirectory(title='Folder to look in', initialdir=self.source.get() or None, mustexist=True)
        if d:
            self.source.set(os.path.normpath(d))

    def _browse_dest(self):
        d = filedialog.askdirectory(title='Folder to copy the statements into', initialdir=self.dest.get() or None)
        if d:
            self.dest.set(os.path.normpath(d))

    def _open(self, path):
        try:
            os.startfile(path)
        except (OSError, AttributeError):
            messagebox.showinfo(APP_NAME, f'Could not open:\n{path}')

    def _open_dest(self):
        if os.path.isdir(self.dest.get()):
            self._open(self.dest.get())

    def _open_report(self):
        if self.last_report and os.path.isfile(self.last_report):
            self._open(self.last_report)

    def _set_running(self, running):
        self.running = running
        self.start_btn.configure(state='disabled' if running else 'normal')
        self.stop_btn.configure(state='normal' if running else 'disabled')

    def _start(self):
        source, dest = self.source.get().strip().strip('"'), self.dest.get().strip().strip('"')
        if not source or not os.path.isdir(source):
            messagebox.showwarning(APP_NAME, 'Choose the folder to look in.')
            return
        if not dest:
            messagebox.showwarning(APP_NAME, 'Choose the folder to copy the statements into.')
            return
        if os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(dest)):
            messagebox.showwarning(APP_NAME, 'The two folders are the same. Choose a different folder to copy into.')
            return
        kinds = tuple(k for k, v in self.kind_vars.items() if v.get())
        if not kinds:
            messagebox.showwarning(APP_NAME, 'Tick at least one file type.')
            return
        if os.path.dirname(os.path.abspath(source).rstrip('\\/')) == os.path.abspath(source).rstrip('\\/') or \
                os.path.splitdrive(os.path.abspath(source))[1] in ('\\', '/'):
            if not messagebox.askokcancel(APP_NAME, 'You chose a whole drive. This can take a long time.\n\nContinue?'):
                return
        self.source.set(source)
        self.dest.set(dest)
        workers = self._safe_int(self.workers, os.cpu_count() or 4, 1, 64)
        timeout = self._safe_int(self.timeout, 60, 10)
        self.workers.set(workers)
        self.timeout.set(timeout)
        self.active_source = source
        self.active_dry = self.dry_var.get()
        self._save_settings()
        opts = {'kinds': kinds, 'include_not_sure': self.not_sure_var.get(),
                'skip_system': self.system_var.get(), 'dry_run': self.active_dry,
                'workers': workers, 'timeout': timeout}
        self.stop.clear()
        self.result = None
        self.counts = {}
        self.sure = self.not_sure = self.copied = 0
        self.last_report = ''
        self.tree.delete(*self.tree.get_children())
        self._log_clear()
        self.report_btn.configure(state='disabled')
        self.open_btn.configure(state='disabled')
        self.bar.configure(mode='indeterminate', value=0)
        self.bar.start(12)
        self.status.set('Starting ...')
        self.counter.set('')
        self._set_running(True)
        self.thread = threading.Thread(target=self._work, args=(source, dest, opts), daemon=True)
        self.thread.start()

    def _work(self, source, dest, opts):
        try:
            core.run(source, dest, opts, self.events.put, self.stop)
        except Exception:
            self.events.put(('error', traceback.format_exc()))

    def _stop(self):
        self.stop.set()
        self.stop_btn.configure(state='disabled')
        self.status.set('Stopping ...')

    def _close(self):
        if self.running:
            if not messagebox.askyesno(APP_NAME, 'A search is still running. Stop it and close?'):
                return
            self.stop.set()
        self._save_settings()
        self.root.destroy()

    # ------------------------------------------------------------------------------------------------- events
    def _poll(self):
        try:
            for _ in range(400):
                self._handle(self.events.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _log_clear(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    def _log(self, text, tag=''):
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n', tag)
        if int(self.log.index('end-1c').split('.')[0]) > 5000:
            self.log.delete('1.0', '1001.0')
        self.log.configure(state='disabled')
        self.log.see('end')

    def _tree_update(self, bank, sure=0, maybe=0):
        c = self.counts.setdefault(bank, [0, 0])
        c[0] += sure
        c[1] += maybe
        if self.tree.exists(bank):
            self.tree.item(bank, values=(bank, c[0], c[1]))
        else:
            self.tree.insert('', 'end', iid=bank, values=(bank, c[0], c[1]))

    def _counter(self):
        parts = [f'Statements: {self.sure:,}', f'Not sure: {self.not_sure:,}']
        if self.active_dry:
            parts.append(f'Would copy: {self.copied:,}')
        else:
            parts.append(f'Copied: {self.copied:,}')
        self.counter.set('   '.join(parts))

    def _handle(self, ev):
        kind = ev[0]
        if kind == 'phase':
            self.status.set(ev[1])
        elif kind == 'walk':
            self.status.set(f'Looking for files ... {ev[1]:,} found so far')
        elif kind == 'total':
            self.bar.stop()
            if ev[1]:
                self.bar.configure(mode='determinate', maximum=ev[1], value=0)
        elif kind == 'progress':
            _k, done, total, name = ev
            self.bar.configure(value=done)
            self.status.set(f'Reading file {done:,} of {total:,}:  {name}')
        elif kind == 'found':
            self._found(ev[1])
        elif kind == 'error':
            self._finish_ui()
            self._log(ev[1], 'error')
            try:
                os.makedirs(SETTINGS_DIR, exist_ok=True)
                with open(CRASH_FILE, 'a', encoding='utf-8') as f:
                    f.write(time.strftime('%Y-%m-%d %H:%M:%S') + '\n' + ev[1] + '\n')
            except OSError:
                pass
            messagebox.showerror(APP_NAME, 'Something went wrong. The details are in the box at the bottom and in\n' + CRASH_FILE)
        elif kind == 'done':
            self._done(ev[1])

    def _found(self, r):
        v = r['verdict']
        try:
            rel = os.path.relpath(r['path'], self.active_source or self.source.get())
        except ValueError:
            rel = r['path']
        act = r.get('action', '')
        if v in ('statement', 'likely'):
            bank = r['bank'] or 'Bank not identified'
            if v == 'statement':
                self.sure += 1
                self._tree_update(bank, sure=1)
            else:
                self.not_sure += 1
                self._tree_update(bank, maybe=1)
            if act in ('copied', 'would copy'):
                self.copied += 1
            tag = 'sure' if v == 'statement' else 'maybe'
            label = bank + (' (card)' if r.get('kind') == 'card' else '') + ('' if v == 'statement' else ' - not sure')
            note = '' if act in ('copied', 'would copy', '') else f'   [{act}]'
            self._log(f'{label:<38} {rel}{note}', tag)
        else:
            self._log(f'{"Could not be read":<38} {rel}   [{r.get("error", "")}]', 'error')
        self._counter()

    def _finish_ui(self):
        self.bar.stop()
        self.bar.configure(mode='determinate')
        self._set_running(False)

    def _done(self, s):
        self._finish_ui()
        self.result = s
        self.last_report = s.get('report', '')
        total_bank = s['statements'] + s['not_sure']
        if s['files'] and not s['stopped']:
            self.bar.configure(value=self.bar['maximum'])
        if s['stopped']:
            head = 'Stopped.'
        else:
            head = 'Finished.'
        if not s['files']:
            self.status.set(f'{head} No files of the chosen types were found in that folder.')
        else:
            verb = 'would be copied' if s['dry_run'] else 'copied'
            self.status.set(f'{head} Looked at {s["read"]:,} of {s["files"]:,} files in {nice_time(s["seconds"])}. '
                            f'{total_bank:,} look like bank statements ({s["statements"]:,} sure, {s["not_sure"]:,} not sure); '
                            f'{s["copied"]:,} {verb}.' + (f' {s["duplicates"]:,} identical files were copied once.' if s['duplicates'] else ''))
        if self.counts:
            ts, tm = sum(c[0] for c in self.counts.values()), sum(c[1] for c in self.counts.values())
            self.tree.insert('', 'end', iid='__total__', values=('All banks', ts, tm), tags=('total',))
            self.tree.tag_configure('total', font=('Segoe UI', 10, 'bold'))
        extra = []
        if s['locked'] or s['scanned']:
            extra.append(f'skipped {s["locked"]:,} password-protected and {s["scanned"]:,} scanned PDFs')
        if s['unreadable']:
            extra.append(f'{s["unreadable"]:,} files could not be read (see the report)')
        if s['folder_problems']:
            extra.append(f'{len(s["folder_problems"]):,} folders or files could not be opened')
        if extra:
            self._log('Also: ' + '; '.join(extra) + '.', 'locked')
        self._counter()
        self.open_btn.configure(state='normal' if (not s['dry_run'] and os.path.isdir(s['dest'])) else 'disabled')
        self.report_btn.configure(state='normal' if self.last_report else 'disabled')


def main():
    mp.freeze_support()
    if len(sys.argv) >= 4 and sys.argv[1] == '--run':
        source, dest = sys.argv[2], sys.argv[3]
        if not os.path.isdir(source) or os.path.normcase(os.path.abspath(source)) == os.path.normcase(os.path.abspath(dest)):
            sys.exit(2)
        core.run(source, dest, {'dry_run': '--scan-only' in sys.argv[4:]}, lambda _ev: None, threading.Event())
        return
    try:                                                              # sharp text on a high-resolution screen
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    try:
        App(root)
        root.mainloop()
    except Exception:
        text = traceback.format_exc()
        try:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            with open(CRASH_FILE, 'a', encoding='utf-8') as f:
                f.write(time.strftime('%Y-%m-%d %H:%M:%S') + '\n' + text + '\n')
        except OSError:
            pass
        messagebox.showerror(APP_NAME, 'The program stopped unexpectedly.\n\n' + text[-900:])


if __name__ == '__main__':
    main()
