"""Tkinter-интерфейс CRF/ФПС."""

import os
import sys
import platform
import subprocess
import threading
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, scrolledtext
import tkinter.font as tkfont

from . import i18n
from .i18n import tr
from .config import load_config, save_config
from .processor import process_document


def resource_path(relative_path: str) -> str:
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).resolve().parent
        candidate = exe_dir / relative_path
        if candidate.exists():
            return str(candidate)
    else:
        candidate = Path(__file__).resolve().parent.parent / relative_path
        if candidate.exists():
            return str(candidate)

    try:
        base = sys._MEIPASS
    except AttributeError:
        base = os.path.abspath(".")
    return os.path.join(base, relative_path)


class CrossReferencerApp(tk.Tk):
    def __init__(self):
        super().__init__()

        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family="Segoe UI", size=9)
        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(family="Segoe UI", size=9)
        style = ttk.Style()
        style.configure(".", font=("Segoe UI", 9))

        cfg = load_config()
        i18n.current_ui_lang = cfg.get('ui_lang', 'ru')
        i18n.current_doc_lang = cfg.get('doc_lang', 'ru')

        self.geometry("960x760")
        self.minsize(720, 560)
        self.input_path = None
        self.output_path = None
        self.last_dir = cfg.get('last_dir', str(Path.home()))
        self.log_history = []

        png_path = resource_path(os.path.join("icons", "window.png"))
        ico_path = resource_path(os.path.join("icons", "window.ico"))
        try:
            if os.path.exists(png_path):
                self._icon_img = tk.PhotoImage(file=png_path)
                self.iconphoto(True, self._icon_img)
            elif os.path.exists(ico_path):
                self.iconbitmap(ico_path)
        except Exception:
            pass

        self._build_ui()
        self._update_ui_text()

        self.log('title')
        self.log('hint_ref')

    def _build_ui(self):
        top_frame = ttk.Frame(self)
        top_frame.pack(fill='x', padx=12, pady=(12, 0))

        self.lbl_title = ttk.Label(top_frame, font=('Segoe UI', 16, 'bold'))
        self.lbl_title.pack(side='left', anchor='w')

        self.doc_lang_frame = ttk.Frame(top_frame)
        self.doc_lang_frame.pack(side='right', padx=(16, 0))
        self.lbl_doc_lang = ttk.Label(self.doc_lang_frame)
        self.lbl_doc_lang.pack(side='left')

        self.doc_lang_var = tk.StringVar(value=i18n.current_doc_lang)
        self.cb_doc_lang = ttk.Combobox(self.doc_lang_frame, textvariable=self.doc_lang_var,
                                        values=['ru', 'en'], state='readonly', width=5)
        self.cb_doc_lang.set(i18n.current_doc_lang)
        self.cb_doc_lang.pack(side='left', padx=(4, 0))
        self.cb_doc_lang.bind('<<ComboboxSelected>>', self._on_doc_lang_change)

        self.ui_lang_frame = ttk.Frame(top_frame)
        self.ui_lang_frame.pack(side='right')
        self.lbl_ui_lang = ttk.Label(self.ui_lang_frame)
        self.lbl_ui_lang.pack(side='left')

        self.ui_lang_var = tk.StringVar(value=i18n.current_ui_lang)
        self.cb_ui_lang = ttk.Combobox(self.ui_lang_frame, textvariable=self.ui_lang_var,
                                       values=['ru', 'en'], state='readonly', width=5)
        self.cb_ui_lang.set(i18n.current_ui_lang)
        self.cb_ui_lang.pack(side='left', padx=(4, 0))
        self.cb_ui_lang.bind('<<ComboboxSelected>>', self._on_ui_lang_change)

        self.lbl_subtitle_desc = ttk.Label(self, foreground='#555', wraplength=930)
        self.lbl_subtitle_desc.pack(anchor='w', padx=12, pady=(0, 10))

        self.ff = ttk.LabelFrame(self)
        self.ff.pack(fill='x', padx=12, pady=6)
        row = ttk.Frame(self.ff); row.pack(fill='x', padx=8, pady=8)

        self.path_var = tk.StringVar()
        self.entry_path = ttk.Entry(row, textvariable=self.path_var, state='readonly')
        self.entry_path.pack(side='left', fill='x', expand=True)

        self.btn_select = ttk.Button(row, command=self.choose_file)
        self.btn_select.pack(side='left', padx=(8, 0))

        self.result_var = tk.StringVar()
        self.lbl_result = ttk.Label(self.ff, textvariable=self.result_var,
                                    foreground='#0a5', wraplength=930)
        self.lbl_result.pack(anchor='w', padx=8, pady=(0, 8))

        actions = ttk.Frame(self); actions.pack(fill='x', padx=12, pady=6)

        self.btn_process = ttk.Button(actions, command=lambda: self._run(False), state='disabled')
        self.btn_process.pack(side='left')

        self.btn_check = ttk.Button(actions, command=lambda: self._run(True), state='disabled')
        self.btn_check.pack(side='left', padx=(8, 0))

        self.btn_open = ttk.Button(actions, command=self.open_folder, state='disabled')
        self.btn_open.pack(side='left', padx=(8, 0))

        self.btn_exit = ttk.Button(actions, command=self.destroy)
        self.btn_exit.pack(side='right')

        self.lf = ttk.LabelFrame(self)
        self.lf.pack(fill='both', expand=True, padx=12, pady=(0, 8))
        self.log_widget = scrolledtext.ScrolledText(
            self.lf, height=20, wrap='word', state='disabled',
            font=('Segoe UI', 9))
        self.log_widget.pack(fill='both', expand=True, padx=6, pady=6)

        hint_frame = ttk.Frame(self)
        hint_frame.pack(fill='x', padx=12, pady=(0, 10))

        self.lbl_hint_f9 = ttk.Label(hint_frame, foreground='#555', wraplength=930)
        self.lbl_hint_f9.pack(anchor='w')

        self.lbl_hint_ref = ttk.Label(hint_frame, foreground='#0a5',
                                      font=('Segoe UI', 9, 'italic'), wraplength=930)
        self.lbl_hint_ref.pack(anchor='w', pady=(2, 0))

    def _update_ui_text(self):
        self.title(tr('title'))
        self.lbl_title.configure(text=tr('app_name'))
        self.lbl_subtitle_desc.configure(text=tr('subtitle'))
        self.ff.configure(text=tr('file_frame'))

        self.lbl_ui_lang.configure(text=tr('lang_label'))
        self.lbl_doc_lang.configure(text=tr('doc_lang_label'))

        if not self.input_path:
            self.path_var.set(tr('no_file'))

        self.btn_select.configure(text=tr('btn_select'))
        self.btn_process.configure(text=tr('btn_process'))
        self.btn_check.configure(text=tr('btn_check'))
        self.btn_open.configure(text=tr('btn_open'))
        self.btn_exit.configure(text=tr('btn_exit'))
        self.lf.configure(text=tr('log_frame'))

        self.lbl_hint_f9.configure(text=tr('hint_f9'))
        self.lbl_hint_ref.configure(text=tr('hint_ref'))

        if self.input_path:
            self.result_var.set("")

    def _on_ui_lang_change(self, event=None):
        i18n.current_ui_lang = self.ui_lang_var.get()
        cfg = load_config()
        cfg['ui_lang'] = i18n.current_ui_lang
        save_config(cfg)
        self._update_ui_text()
        self._redraw_log()

    def _on_doc_lang_change(self, event=None):
        i18n.current_doc_lang = self.doc_lang_var.get()
        cfg = load_config()
        cfg['doc_lang'] = i18n.current_doc_lang
        save_config(cfg)

    def log(self, key=None, raw_text=None, **kwargs):
        if key:
            item = ('key', key, kwargs)
        else:
            item = ('raw', raw_text, {})
        self.log_history.append(item)
        self.after(0, self._append_log_item, item)

    def _append_log_item(self, item):
        item_type, val, kwargs = item
        msg = tr(val, **kwargs) if item_type == 'key' else val
        self.log_widget.configure(state='normal')
        self.log_widget.insert('end', msg + '\n')
        self.log_widget.see('end')
        self.log_widget.configure(state='disabled')

    def _redraw_log(self):
        self.log_widget.configure(state='normal')
        self.log_widget.delete('1.0', 'end')
        for item_type, val, kwargs in self.log_history:
            msg = tr(val, **kwargs) if item_type == 'key' else val
            self.log_widget.insert('end', msg + '\n')
        self.log_widget.see('end')
        self.log_widget.configure(state='disabled')

    def choose_file(self):
        path = filedialog.askopenfilename(
            title="Select .docx" if i18n.current_ui_lang == 'en' else "Выберите .docx",
            initialdir=self.last_dir,
            filetypes=[("Word", "*.docx"), ("All", "*.*")])
        if not path:
            return
        if not path.lower().endswith('.docx'):
            messagebox.showwarning(
                tr('msg_wrong_format_title'), tr('msg_wrong_format_text'))
            return
        self.input_path = path
        self.last_dir = str(Path(path).parent)

        cfg = load_config()
        cfg['last_dir'] = self.last_dir
        save_config(cfg)

        self.path_var.set(path)
        self.result_var.set("")
        self.btn_process.configure(state='normal')
        self.btn_check.configure(state='normal')
        self.btn_open.configure(state='disabled')
        self.log('log_file_selected', path=path)

    def _run(self, dry_run):
        if not self.input_path:
            return
        if dry_run:
            output = self.input_path
            self.log(raw_text="=" * 60)
            self.log('log_check_mode')
        else:
            output = self._unique_output_path(self.input_path)
            self.output_path = output
            self.log(raw_text="=" * 60)
            self.log('log_result', path=output)

        self.btn_process.configure(state='disabled')
        self.btn_check.configure(state='disabled')
        self.btn_open.configure(state='disabled')
        self.result_var.set("")

        def worker():
            try:
                stats = process_document(
                    self.input_path, output, log=self.log, dry_run=dry_run)
                self.after(0, self._on_success, output, stats, dry_run)
            except Exception as e:
                self.after(0, self._on_error, e)
        threading.Thread(target=worker, daemon=True).start()

    def _on_success(self, output, stats, dry_run):
        if not dry_run:
            self.result_var.set(tr('msg_done_path', path=output))
            self.btn_open.configure(state='normal')
        else:
            self.result_var.set(tr('msg_check_done'))

        self.log(raw_text="=" * 60)
        self.log('log_done')
        self.log('hint_ref')

        self.btn_process.configure(state='normal')
        self.btn_check.configure(state='normal')

        msg = tr('msg_stats', **stats)
        title = tr('msg_check') if dry_run else tr('msg_success')
        messagebox.showinfo(title, msg)

    def _on_error(self, exc):
        self.log('log_error', exc=exc)
        self.btn_process.configure(state='normal')
        self.btn_check.configure(state='normal')

        msg = str(exc)
        if isinstance(exc, PermissionError):
            msg += tr('msg_error_word')
        messagebox.showerror(tr('msg_error_title'), msg)

    def open_folder(self):
        if not self.output_path:
            return
        folder = str(Path(self.output_path).parent)
        try:
            if platform.system() == 'Windows':
                os.startfile(folder)
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', folder])
            else:
                subprocess.Popen(['xdg-open', folder])
        except Exception as e:
            self.log('msg_folder_err', e=e)

    @staticmethod
    def _unique_output_path(input_path):
        p = Path(input_path)
        suffix = "_autonumbering" if i18n.current_doc_lang == 'en' else "_автонумерация"
        c = p.with_name(f"{p.stem}{suffix}{p.suffix}")
        if not c.exists():
            return str(c)
        i = 1
        while True:
            c = p.with_name(f"{p.stem}{suffix}_{i}{p.suffix}")
            if not c.exists():
                return str(c)
            i += 1