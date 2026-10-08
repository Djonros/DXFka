# -*- coding: utf-8 -*-
"""Главное окно приложения."""

import logging
import os
import queue
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional, Tuple

from dxfka.config import (
    APP_TITLE, APP_VERSION, DEFAULT_OUT_SUBDIR, LOG_FILE,
    STARTUP_INSTRUCTIONS,
    SUPPORT_EMAIL, TRIAL_CORRUPTED_MSG, TRIAL_EXHAUSTED_MSG,
    TRIAL_EXPORT_LIMIT, VENDOR)
from dxfka.kompas.connection import clear_com_cache
from dxfka.licensing import LicenseManager
from dxfka.models import PartInfo, part_file_label, part_type_text
from dxfka.ui import theme
from dxfka.ui.widgets import add_tooltip
from dxfka.utils import app_dir, load_settings, save_settings
from dxfka.worker import ExportWorker

# Колонки таблицы деталей: (id, заголовок, ширина, растягивается).
PART_COLUMNS = [("file", "Файл", 170, False), ("type", "Тип", 105, False),
                ("mark", "Обозначение", 140, False),
                ("name", "Наименование", 200, True),
                ("thk", "Толщина", 95, False), ("status", "Статус", 210, True)]
CHECK_ON, CHECK_OFF = "☑", "☐"
STATUS_TEXT = {"OK": "✓ Готово", "ERROR": "⚠ Ошибка", "SKIP": "Пропущена"}


class App(tk.Tk):
    """Главное окно приложения."""

    def __init__(self, lic: LicenseManager, logger: logging.Logger):
        super().__init__()
        self.lic = lic
        self.logger = logger

        self.title(f"{APP_TITLE} {APP_VERSION} — {VENDOR}")
        self.geometry("1080x760")
        self.minsize(900, 640)

        # Мосты GUI <-> рабочий поток.
        self.msg_queue: "queue.Queue" = queue.Queue()
        self.resume_event = threading.Event()
        self.cancel_event = threading.Event()
        self.skip_event = threading.Event()
        self.worker: Optional[ExportWorker] = None

        # Журнал: в главном окне не показываем — окно по запросу из меню.
        self.log_text: Optional[tk.Text] = None
        self._log_window: Optional[tk.Toplevel] = None
        self._log_buffer: List[Tuple[str, str]] = []

        # Данные списка деталей.
        self.part_vars: List[Tuple[tk.BooleanVar, PartInfo]] = []
        self.rows: Dict[str, str] = {}          # uid детали -> строка таблицы
        self._row_state: Dict[str, str] = {}    # uid -> OK/ERROR/SKIP/RUN/""
        self._row_vars: Dict[str, tk.BooleanVar] = {}
        self._row_parts: Dict[str, PartInfo] = {}

        self._apply_theme()
        self._build_ui()
        self._build_menu()
        self._set_buttons("idle")
        self._update_license_label()
        self._log_gui("info", STARTUP_INSTRUCTIONS)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    # Тема и меню
    # ------------------------------------------------------------------

    def _apply_theme(self) -> None:
        """Единый светлый стиль (см. dxfka/ui/theme.py)."""
        theme.apply_theme(self)

    def _build_menu(self) -> None:
        """Строка меню: Файл / Настройки / Справка (те же обработчики)."""
        menubar = tk.Menu(self)

        m_file = tk.Menu(menubar, tearoff=0)
        m_file.add_command(label="Выбрать сборку…",
                           command=self.on_browse_assembly)
        m_file.add_command(label="Добавить файлы деталей…",
                           command=self.on_add_manual)
        m_file.add_separator()
        m_file.add_command(label="Показать журнал",
                           command=self.show_log_window)
        m_file.add_separator()
        m_file.add_command(label="Выход", command=self.on_close)
        menubar.add_cascade(label="Файл", menu=m_file)

        m_settings = tk.Menu(menubar, tearoff=0)
        m_settings.add_command(label="Папка DXF…",
                               command=self.on_browse_out)
        m_settings.add_command(label="Шаблон фрагмента…",
                               command=self.on_browse_template)
        m_settings.add_command(label="Сбросить шаблон",
                               command=self.on_clear_template)
        m_settings.add_separator()
        m_settings.add_command(label="Сброс кеша COM…",
                               command=self.on_clear_com_cache)
        menubar.add_cascade(label="Настройки", menu=m_settings)

        m_mode = tk.Menu(menubar, tearoff=0)
        m_mode.add_radiobutton(label="Автоматически", value="auto",
                               variable=self.mode_var,
                               command=self.on_mode_changed)
        m_mode.add_radiobutton(label="Вручную", value="manual",
                               variable=self.mode_var,
                               command=self.on_mode_changed)
        menubar.add_cascade(label="Обработка", menu=m_mode)
        self._menubar = menubar

        m_help = tk.Menu(menubar, tearoff=0)
        m_help.add_command(label="Инструкция",
                           command=self.show_instructions)
        m_help.add_command(label="Лицензия…",
                           command=self.show_license_dialog)
        m_help.add_command(label="О программе", command=self.show_about)
        menubar.add_cascade(label="Справка", menu=m_help)
        self._mode_menu_index = menubar.index("Обработка")

        self.configure(menu=menubar)

    # ------------------------------------------------------------------
    # Построение интерфейса
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 16}

        # ---- Подвал: кнопки действий, видны ВСЕГДА ----
        # Пакуется первым со стороны bottom: место резервируется у нижнего
        # края, и при нехватке высоты сжимается таблица, а не кнопки.
        frm_btn = ttk.Frame(self, padding=(16, 10, 16, 14))
        frm_btn.pack(side="bottom", fill="x")
        self.btn_export = ttk.Button(frm_btn, text="Начать экспорт",
                                     style="Accent.TButton",
                                     command=self.on_start_export)
        self.btn_export.pack(side="left")
        self.btn_continue = ttk.Button(frm_btn, text="Продолжить",
                                       command=self.on_continue)
        self.btn_continue.pack(side="left", padx=(8, 0))
        self.btn_skip = ttk.Button(frm_btn, text="Пропустить деталь",
                                   command=self.on_skip)
        self.btn_skip.pack(side="left", padx=(8, 0))
        self.btn_log = ttk.Button(frm_btn, text="Журнал",
                                  command=self.show_log_window)
        self.btn_log.pack(side="right")
        self.btn_cancel = ttk.Button(frm_btn, text="Остановить",
                                     command=self.on_cancel)
        self.btn_cancel.pack(side="right", padx=(0, 8))

        # ---- Шапка: название и лицензия ----
        hdr_o, hdr = theme.card(self)
        hdr_o.pack(fill="x", pady=(14, 0), **pad)
        logo = tk.Canvas(hdr, width=38, height=38, bg=theme.CARD,
                         highlightthickness=0)
        logo.create_rectangle(1, 1, 37, 37, fill=theme.ACCENT,
                              outline=theme.ACCENT)
        logo.create_text(19, 19, text="D", fill="white",
                         font=theme.font(17, True))
        logo.pack(side="left", padx=(0, 12))
        titles = ttk.Frame(hdr, style="Card.TFrame")
        titles.pack(side="left")
        ttk.Label(titles, text=f"DXFka {APP_VERSION}",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(titles, text="Пакетный экспорт DXF из сборок КОМПАС-3D v24",
                  style="Muted.TLabel").pack(anchor="w")
        right = ttk.Frame(hdr, style="Card.TFrame")
        right.pack(side="right")
        self.license_var = tk.StringVar(value="")
        self.license_label = tk.Label(right, textvariable=self.license_var,
                                      font=theme.font(9, True), padx=10,
                                      pady=4, cursor="hand2")
        self.license_label.pack(anchor="e")
        self.license_label.bind("<Button-1>",
                                lambda _e: self.show_license_dialog())
        ttk.Label(right, text=f"{VENDOR} · {SUPPORT_EMAIL}",
                  style="Muted.TLabel").pack(anchor="e", pady=(4, 0))

        # ---- Исходные данные ----
        src_o, frm_top = theme.card(self)
        src_o.pack(fill="x", pady=(12, 0), **pad)
        ttk.Label(frm_top, text="Исходные данные",
                  style="Section.TLabel").grid(row=0, column=0, columnspan=4,
                                               sticky="w", pady=(0, 8))

        self.asm_var = tk.StringVar()
        settings = load_settings()
        self.settings = settings
        self.out_var = tk.StringVar(
            value=settings.get("out_dir")
            or os.path.join(app_dir(), DEFAULT_OUT_SUBDIR))
        self.tpl_var = tk.StringVar(value=settings.get("template", ""))
        self.mode_var = tk.StringVar(
            value="auto" if settings.get("mode") == "auto" else "manual")

        def field(row: int, label: str, var: tk.StringVar) -> ttk.Entry:
            ttk.Label(frm_top, text=label, style="Card.TLabel").grid(
                row=row, column=0, sticky="w", padx=(0, 10), pady=3)
            entry = ttk.Entry(frm_top, textvariable=var)
            entry.grid(row=row, column=1, sticky="we", pady=3)
            return entry

        field(1, "Сборка (.a3d)", self.asm_var)
        self.btn_browse_asm = ttk.Button(
            frm_top, text="Обзор…", style="Ghost.TButton",
            command=self.on_browse_assembly)
        self.btn_browse_asm.grid(row=1, column=2, padx=8, pady=3)
        self.btn_load = ttk.Button(
            frm_top, text="Загрузить детали", style="Ghost.TButton",
            command=self.on_load_parts)
        self.btn_load.grid(row=1, column=3, sticky="we", pady=3)

        field(2, "Папка для DXF", self.out_var)
        self.btn_browse_out = ttk.Button(
            frm_top, text="Обзор…", style="Ghost.TButton",
            command=self.on_browse_out)
        self.btn_browse_out.grid(row=2, column=2, padx=8, pady=3)

        self.tpl_entry = field(3, "Шаблон фрагмента", self.tpl_var)
        self.btn_browse_tpl = ttk.Button(
            frm_top, text="Обзор…", style="Ghost.TButton",
            command=self.on_browse_template)
        self.btn_browse_tpl.grid(row=3, column=2, padx=8, pady=3)
        self.btn_clear_tpl = ttk.Button(
            frm_top, text="Сброс", style="Ghost.TButton",
            command=self.on_clear_template)
        self.btn_clear_tpl.grid(row=3, column=3, sticky="we", pady=3)
        frm_top.columnconfigure(1, weight=1)

        # ---- Ход работы (над таблицей: инструкция паузы должна быть
        # видна целиком, таблица забирает остаток места) ----
        pr_o, frm_status = theme.card(self)
        pr_o.pack(fill="x", pady=(12, 0), **pad)
        line = ttk.Frame(frm_status, style="Card.TFrame")
        line.pack(fill="x")
        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(line, textvariable=self.status_var,
                  style="Section.TLabel").pack(side="left")
        self.step_var = tk.StringVar(value=self._mode_text())
        ttk.Label(line, textvariable=self.step_var,
                  style="Muted.TLabel").pack(side="right")
        self.progress = ttk.Progressbar(
            frm_status, mode="determinate",
            style="Thin.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(8, 0))
        self.instr_var = tk.StringVar(value="")
        self.instr_label = ttk.Label(frm_status, textvariable=self.instr_var,
                                     justify="left", wraplength=960,
                                     style="Instr.TLabel")
        self.instr_label.pack(anchor="w", pady=(6, 0))
        frm_status.bind(
            "<Configure>",
            lambda e: self.instr_label.configure(wraplength=e.width - 24))

        # ---- Детали ----
        pl_o, frm_parts = theme.card(self)
        pl_o.pack(fill="both", expand=True, pady=(12, 0), **pad)
        toolbar = ttk.Frame(frm_parts, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(0, 8))
        ttk.Label(toolbar, text="Детали",
                  style="Section.TLabel").pack(side="left")
        self.counts_var = tk.StringVar(value="")
        self.counts_label = ttk.Label(toolbar, textvariable=self.counts_var,
                                      style="Muted.TLabel")
        self.counts_label.pack(side="left", padx=10)
        self.btn_clear_all = ttk.Button(
            toolbar, text="Снять все", style="Ghost.TButton",
            command=lambda: self._toggle_all(False))
        self.btn_clear_all.pack(side="right")
        self.btn_select_all = ttk.Button(
            toolbar, text="Выбрать все", style="Ghost.TButton",
            command=lambda: self._toggle_all(True))
        self.btn_select_all.pack(side="right", padx=(0, 6))
        self.btn_add = ttk.Button(toolbar, text="Добавить файлы",
                                  style="Ghost.TButton",
                                  command=self.on_add_manual)
        self.btn_add.pack(side="right", padx=(0, 6))

        container = ttk.Frame(frm_parts, style="Card.TFrame")
        container.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(
            container, columns=[c[0] for c in PART_COLUMNS],
            show="tree headings", style="Parts.Treeview", selectmode="none")
        self.tree.heading("#0", text="")
        self.tree.column("#0", width=36, minwidth=36, stretch=False,
                         anchor="center")
        for cid, title, width, stretch in PART_COLUMNS:
            self.tree.heading(cid, text=title, anchor="w")
            self.tree.column(cid, width=width, minwidth=60, stretch=stretch,
                             anchor="w")
        scroll = ttk.Scrollbar(container, orient="vertical",
                               command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("even", background=theme.CARD)
        self.tree.tag_configure("odd", background=theme.ROW_ALT)
        for state, color in theme.STATUS_COLORS.items():
            self.tree.tag_configure(state, foreground=color)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<space>", self._on_tree_space)
        self.tree.bind("<Motion>", self._on_tree_motion)
        self.tree.bind("<Leave>", lambda _e: self._hide_row_tip())
        self._row_tip: Optional[tk.Toplevel] = None
        self._row_tip_iid = ""
        self._update_counts()

        # ---- Всплывающие подсказки ----
        for widget, text in (
            (self.btn_browse_asm, "Выбрать файл сборки КОМПАС (.a3d)"),
            (self.btn_load,
             "Проанализировать сборку: найти все детали, включая\n"
             "исполнения и локальные (встроенные в сборку)"),
            (self.btn_browse_out,
             "Папка, в которую будут сохранены DXF-файлы"),
            (self.btn_browse_tpl,
             "Шаблон фрагмента .frw (необязательно): фрагменты\n"
             "создаются из него — надписи не придётся удалять вручную"),
            (self.btn_clear_tpl, "Убрать шаблон — используется пустой фрагмент"),
            (self.btn_add,
             "Добавить файлы деталей (.m3d) вручную, если обход\n"
             "сборки не нашёл нужные детали"),
            (self.btn_select_all, "Отметить все детали в списке"),
            (self.btn_clear_all, "Снять отметки со всех деталей"),
            (self.counts_label,
             "Отметка детали — щелчок по флажку в первой колонке;\n"
             "подробности о детали — при наведении на строку"),
            (self.btn_export, "Экспортировать отмеченные детали в DXF"),
            (self.btn_continue,
             "Продолжить после паузы\n(проверьте вид/развертку в КОМПАС)"),
            (self.btn_skip, "Пропустить текущую деталь без сохранения DXF"),
            (self.btn_cancel,
             "Остановить выполнение\n(КОМПАС завершит текущий шаг)"),
            (self.btn_log, "Подробный журнал работы"),
            (self.license_label,
             "Статус лицензии — нажмите для подробностей"),
        ):
            add_tooltip(widget, text)

    # ------------------------------------------------------------------
    # Управление состоянием кнопок
    # ------------------------------------------------------------------

    def _set_buttons(self, state: str) -> None:
        """idle — простой; working — идёт работа; paused — ждём пользователя."""
        idle = state == "idle"
        working = state == "working"
        paused = state == "paused"
        for btn in (self.btn_load, self.btn_add, self.btn_browse_asm,
                    self.btn_browse_out, self.btn_browse_tpl,
                    self.btn_clear_tpl, self.btn_select_all,
                    self.btn_clear_all, self.btn_export):
            btn.configure(state="normal" if idle else "disabled")
        self.btn_cancel.configure(
            state="normal" if (working or paused) else "disabled")
        self.btn_continue.configure(state="normal" if paused else "disabled")
        self.btn_skip.configure(state="normal" if paused else "disabled")
        menubar = getattr(self, "_menubar", None)
        if menubar is not None:
            # Режим меняется только между запусками.
            menubar.entryconfigure(self._mode_menu_index,
                                   state="normal" if idle else "disabled")

    def _worker_alive(self) -> bool:
        return bool(self.worker and self.worker.is_alive())

    # ------------------------------------------------------------------
    # Журнал и статус
    # ------------------------------------------------------------------

    def _log_gui(self, level: str, text: str) -> None:
        """Буфер журнала + вывод в окно журнала, если оно открыто."""
        tag = level if level in ("info", "warning", "error", "ok") else "info"
        for line in text.splitlines() or [""]:
            self._log_buffer.append((tag, line))
        if len(self._log_buffer) > 5000:
            del self._log_buffer[:-5000]
        widget = self.log_text
        if widget is None:
            return
        try:
            if not widget.winfo_exists():
                self.log_text = None
                return
            widget.configure(state="normal")
            widget.insert("end", text + "\n", tag)
            widget.see("end")
            widget.configure(state="disabled")
        except Exception:
            self.log_text = None

    def show_log_window(self) -> None:
        """Окно журнала (меню «Файл → Показать журнал»)."""
        if (self._log_window is not None
                and self._log_window.winfo_exists()):
            self._log_window.deiconify()
            self._log_window.lift()
            return
        win = tk.Toplevel(self)
        win.configure(bg=theme.BG)
        self._log_window = win
        win.title(f"Журнал — {APP_TITLE}")
        win.geometry("900x420")

        frm = ttk.Frame(win, padding=(8, 8))
        frm.pack(fill="both", expand=True)
        self.log_text = tk.Text(frm, state="disabled", wrap="none",
                                font=("Consolas", 9))
        scroll = ttk.Scrollbar(frm, orient="vertical",
                               command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        for tag, color in (("info", "#202020"), ("warning", "#b36b00"),
                           ("error", "#b00020"), ("ok", "#0a7a30")):
            self.log_text.tag_configure(tag, foreground=color)
        self.log_text.configure(state="normal")
        for tag, line in self._log_buffer:
            self.log_text.insert("end", line + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

        btns = ttk.Frame(win, padding=(8, 0, 8, 8))
        btns.pack(fill="x")

        def open_log_file() -> None:
            try:
                os.startfile(LOG_FILE)      # noqa: S606 — файл журнала рядом с программой
            except Exception as exc:
                messagebox.showerror(APP_TITLE,
                                     f"Не удалось открыть журнал: {exc}",
                                     parent=win)

        ttk.Button(btns, text="Открыть файл журнала",
                   command=open_log_file).pack(side="left")
        ttk.Button(btns, text="Закрыть",
                   command=win.destroy).pack(side="right")

    def _update_license_label(self) -> None:
        self.license_var.set(self.lic.status_text())
        if self.lic.licensed:
            fg, bg = theme.OK, theme.OK_BG
        elif self.lic.corrupted:
            fg, bg = theme.ERR, theme.ERR_BG
        else:
            fg, bg = theme.WARN, theme.WARN_BG
        self.license_label.configure(foreground=fg, background=bg)

    # ------------------------------------------------------------------
    # Выбор путей
    # ------------------------------------------------------------------

    def on_browse_assembly(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите файл сборки КОМПАС",
            filetypes=[("Сборка КОМПАС (*.a3d)", "*.a3d"),
                       ("Все файлы", "*.*")])
        if path:
            self.asm_var.set(path)

    def on_browse_out(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку для DXF-файлов")
        if path:
            self.out_var.set(path)
            self._save_settings()

    def on_browse_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите шаблон фрагмента КОМПАС",
            filetypes=[("Шаблон фрагмента (*.frw)", "*.frw"),
                         ("Все файлы", "*.*")])
        if path:
            self.tpl_var.set(path)
            self._save_settings()
            self._log_gui("info",
                          f"Шаблон фрагмента: {path}\n"
                          "(фрагменты будут создаваться из шаблона)")

    def on_clear_template(self) -> None:
        self.tpl_var.set("")
        self._save_settings()
        self._log_gui("info", "Шаблон фрагмента сброшен — "
                              "используется пустой фрагмент")

    def _save_settings(self) -> None:
        """Сохранить текущие настройки GUI (шаблон, папку вывода)."""
        self.settings["template"] = self.tpl_var.get().strip()
        self.settings["out_dir"] = self.out_var.get().strip()
        self.settings["mode"] = self.mode_var.get()
        save_settings(self.settings)

    def _mode_text(self) -> str:
        return ("Режим: автоматически" if self.mode_var.get() == "auto"
                else "Режим: вручную")

    def on_mode_changed(self) -> None:
        """Меню «Обработка»: автоматически / вручную."""
        self._save_settings()
        self.step_var.set(self._mode_text())
        if self.mode_var.get() == "auto":
            self._log_gui(
                "info",
                "Режим «Автоматически»: толщина листовых деталей берётся "
                "из модели, паузы пропускаются. Если толщину определить "
                "не удалось (нелистовая деталь) — деталь обрабатывается "
                "вручную.")
        else:
            self._log_gui("info", "Режим «Вручную»: пауза на каждой детали.")

    def on_clear_com_cache(self) -> None:
        """Сброс кеша COM-обёрток pywin32 (кнопка «Сброс кеша COM»)."""
        if self._worker_alive():
            messagebox.showwarning(
                APP_TITLE, "Дождитесь окончания экспорта — "
                           "кеш нельзя сбрасывать во время работы.")
            return
        if not messagebox.askyesno(
                APP_TITLE,
                "Удалить кеш COM-обёрток pywin32 (gen_py)?\n\n"
                "Помогает при ошибках вида\n"
                "«win32com.gen_py has no attribute …»\n"
                "после обновления КОМПАС-3D.\n"
                "Кеш будет заново создан при следующем экспорте."):
            return
        removed = clear_com_cache()
        if removed:
            self._log_gui("ok", "Кеш COM сброшен:\n" + "\n".join(removed))
        else:
            self._log_gui("info", "Кеш COM на диске не найден — "
                                  "выгружены только модули из памяти")
        self._log_gui("info", "Заново загрузите детали и повторите "
                              "экспорт. Если ошибки сохранятся — "
                              "перезапустите приложение.")

    # ------------------------------------------------------------------
    # Загрузка деталей из сборки / вручную
    # ------------------------------------------------------------------

    def on_load_parts(self) -> None:
        path = self.asm_var.get().strip()
        if not path:
            messagebox.showwarning(APP_TITLE, "Укажите файл сборки (.a3d).")
            return
        if not os.path.isfile(path):
            messagebox.showwarning(APP_TITLE,
                                   f"Файл не найден:\n{path}")
            return
        self._log_gui("info", f"Анализ сборки: {path}")
        self.logger.info(f"Анализ сборки: {path}")
        self._start_worker("traverse", os.path.abspath(path))

    def on_add_manual(self) -> None:
        """Ручное добавление файлов деталей (если обход не удался)."""
        paths = filedialog.askopenfilenames(
            title="Добавьте файлы деталей (.m3d)",
            filetypes=[("Детали КОМПАС (*.m3d)", "*.m3d"),
                       ("Все файлы", "*.*")])
        if not paths:
            return
        existing = {os.path.normcase(os.path.abspath(p.file_path))
                    for _, p in self.part_vars}
        added = 0
        for path in paths:
            key = os.path.normcase(os.path.abspath(path))
            if key in existing:
                self._log_gui("warning",
                              f"Уже в списке (пропущено): {os.path.basename(path)}")
                continue
            existing.add(key)
            if os.path.splitext(path)[1].lower() != ".m3d":
                self._log_gui("warning",
                              f"Не .m3d — пропущено: {os.path.basename(path)}")
                continue
            stem = os.path.splitext(os.path.basename(path))[0]
            part = PartInfo(file_path=os.path.abspath(path),
                            name=stem, is_sheet=None)
            self._append_part_row(part, selected=True)
            added += 1
        if added:
            self._log_gui("info", f"Вручную добавлено деталей: {added}")
        self._update_counts()

    # ------------------------------------------------------------------
    # Список деталей
    # ------------------------------------------------------------------

    def _rebuild_part_list(self, parts: List[PartInfo]) -> None:
        """Полная перестройка списка после обхода сборки."""
        self._hide_row_tip()
        self.tree.delete(*self.tree.get_children())
        self.rows.clear()
        self.part_vars.clear()
        self._row_state.clear()
        self._row_vars.clear()
        self._row_parts.clear()
        ordered = sorted(parts, key=lambda p: (p.marking, p.name, p.file_path))
        for part in ordered:
            # Нелистовые детали (по эвристике материала) — сняты с выбора;
            # неопределённый тип остаётся выбранным.
            self._append_part_row(part, selected=part.is_sheet is not False)

    def _append_part_row(self, part: PartInfo, selected: bool = True) -> None:
        """Добавление одной строки в таблицу деталей."""
        var = tk.BooleanVar(value=selected)
        marking = part.marking
        if part.performance:
            marking = f"{marking} (исп. {part.performance})".strip()
        iid = self.tree.insert(
            "", "end", text="",
            values=(part_file_label(part), part_type_text(part.is_sheet),
                    marking, part.name, "—", "Ожидает"))
        self.rows[part.uid] = iid
        self._row_vars[iid] = var
        self._row_parts[iid] = part
        self._row_state[iid] = ""
        self.part_vars.append((var, part))
        self._refresh_row(iid)

    def _refresh_row(self, iid: str) -> None:
        """Флажок, цвет и чередование строки по её состоянию."""
        checked = self._row_vars[iid].get()
        state = self._row_state.get(iid, "")
        index = self.tree.index(iid)
        tags = ["odd" if index % 2 else "even"]
        if state:
            tags.append(state)
        elif not checked:
            tags.append("OFF")
        self.tree.item(iid, text=CHECK_ON if checked else CHECK_OFF,
                       tags=tags)
        if not state:
            self.tree.set(iid, "status", "Ожидает" if checked else "Снята")

    def _toggle_row(self, iid: str) -> None:
        if not iid or iid not in self._row_vars or self._worker_alive():
            return
        var = self._row_vars[iid]
        var.set(not var.get())
        self._refresh_row(iid)
        self._update_counts()

    def _on_tree_click(self, event: tk.Event) -> Optional[str]:
        """Щелчок по флажку (первая колонка) — отметить/снять деталь."""
        if self.tree.identify_region(event.x, event.y) == "heading":
            return None
        if self.tree.identify_column(event.x) == "#0":
            self._toggle_row(self.tree.identify_row(event.y))
            return "break"
        return None

    def _on_tree_space(self, _event: tk.Event) -> str:
        self._toggle_row(self.tree.focus())
        return "break"

    def _row_details(self, part: PartInfo) -> str:
        details = (f"Файл: {part.file_path}\n" if part.file_path
                   else "Локальная деталь (встроена в сборку)\n")
        details += f"Обозначение: {part.marking}\nНаименование: {part.name}"
        if part.material:
            details += f"\nМатериал: {part.material}"
        if part.performance:
            details += f"\nИсполнение: {part.performance}"
        if part.message:
            details += f"\n\n{part.message}"
        return details

    def _on_tree_motion(self, event: tk.Event) -> None:
        """Подсказка с подробностями о детали при наведении на строку."""
        iid = self.tree.identify_row(event.y)
        if iid == self._row_tip_iid:
            return
        self._hide_row_tip()
        part = self._row_parts.get(iid)
        if part is None:
            return
        self._row_tip_iid = iid
        tip = tk.Toplevel(self)
        tip.wm_overrideredirect(True)
        tip.wm_geometry(f"+{event.x_root + 16}+{event.y_root + 18}")
        tk.Label(tip, text=self._row_details(part), justify="left",
                 background="#ffffff", foreground=theme.TXT,
                 relief="solid", borderwidth=1, font=theme.font(9),
                 wraplength=460, padx=8, pady=6).pack()
        self._row_tip = tip

    def _hide_row_tip(self) -> None:
        self._row_tip_iid = ""
        if self._row_tip is not None:
            try:
                self._row_tip.destroy()
            except Exception:
                pass
            self._row_tip = None

    def _toggle_all(self, value: bool) -> None:
        for iid, var in self._row_vars.items():
            var.set(value)
            self._refresh_row(iid)
        self._update_counts()

    def _update_counts(self) -> None:
        total = len(self.part_vars)
        chosen = sum(1 for var, _ in self.part_vars if var.get())
        done = sum(1 for st in self._row_state.values() if st == "OK")
        text = f"найдено {total} · выбрано {chosen}"
        if done:
            text += f" · готово {done}"
        self.counts_var.set(text)

    def _mark_part_running(self, uid: str) -> None:
        iid = self.rows.get(uid)
        if iid is None:
            return
        self._row_state[iid] = "RUN"
        self.tree.set(iid, "status", "● Экспорт…")
        self._refresh_row(iid)
        self.tree.see(iid)

    def _update_part_row(self, uid: str, status: str,
                         message: str, is_sheet: Optional[bool],
                         thickness: str = "") -> None:
        iid = self.rows.get(uid)
        if iid is None:
            return
        self._row_state[iid] = status if status in STATUS_TEXT else ""
        text = STATUS_TEXT.get(status, "—")
        if status != "OK" and message:
            text += f": {message}"
        self.tree.set(iid, "status", text[:120])
        self._row_parts[iid].message = message
        if thickness:
            self.tree.set(iid, "thk", f"{thickness} мм")
        if is_sheet is not None:
            self.tree.set(iid, "type", part_type_text(is_sheet))
        self._refresh_row(iid)
        self._update_counts()

    # ------------------------------------------------------------------
    # Запуск фоновых операций
    # ------------------------------------------------------------------

    def _start_worker(self, mode: str, payload: Any) -> None:
        self.cancel_event.clear()
        self.resume_event.clear()
        self.skip_event.clear()
        self._save_settings()
        self.worker = ExportWorker(self.msg_queue, self.resume_event,
                                   self.cancel_event, self.skip_event,
                                   mode, payload, self.lic, self.logger)
        if mode == "export":
            self.worker.template_path = self.tpl_var.get().strip() or None
            self.worker.auto_mode = self.mode_var.get() == "auto"
        self.worker.start()
        self._set_buttons("working")

    def on_start_export(self) -> None:
        selected = [p for var, p in self.part_vars if var.get()]
        if not selected:
            messagebox.showwarning(APP_TITLE,
                                   "Отметьте хотя бы одну деталь в списке.")
            return
        out_dir = self.out_var.get().strip()
        if not out_dir:
            messagebox.showwarning(APP_TITLE, "Укажите папку для DXF.")
            return
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_TITLE,
                                 f"Не удалось создать папку:\n{exc}")
            return

        # Проверки пробной версии/лицензии ДО старта.
        if not self.lic.licensed:
            if self.lic.corrupted:
                messagebox.showerror(APP_TITLE, TRIAL_CORRUPTED_MSG)
                self.show_license_dialog()
                return
            remaining = self.lic.remaining
            if remaining < 1:
                messagebox.showerror(APP_TITLE, TRIAL_EXHAUSTED_MSG)
                self.show_license_dialog()
                return
            if len(selected) > remaining:
                if not messagebox.askyesno(
                        APP_TITLE,
                        f"Выбрано деталей: {len(selected)}, но в пробной версии "
                        f"осталось экспортов: {remaining}.\n"
                        f"Будут экспортированы только первые {remaining}. "
                        "Продолжить?"):
                    return

        # Пути DXF считаются в рабочем потоке: толщина запрашивается
        # перед каждой деталью, DXF кладётся в подпапку «толщина».
        self._log_gui(
            "info",
            f"Начинаю экспорт: {len(selected)} дет., папка: {out_dir}\n"
            "Для каждой детали будет запрошена толщина (мм) — DXF "
            "сохраняется в подпапку с этим номером.")
        self.progress.configure(value=0, maximum=len(selected))
        self.step_var.set(f"готово 0 из {len(selected)}")
        for iid in self._row_vars:
            if self._row_state.get(iid) in ("OK", "ERROR", "SKIP", "RUN"):
                self._row_state[iid] = ""
                self.tree.set(iid, "thk", "—")
            self._refresh_row(iid)
        self._update_counts()
        self._start_worker("export", (out_dir, selected))

    # ------------------------------------------------------------------
    # Кнопки паузы/отмены
    # ------------------------------------------------------------------

    def on_continue(self) -> None:
        """«Продолжить»: разбудить рабочий поток."""
        self.resume_event.set()
        self._set_buttons("working")
        self.instr_var.set("")
        self.status_var.set("Продолжаю…")

    def on_skip(self) -> None:
        """«Пропустить деталь»: пропустить текущую и идти дальше."""
        self.skip_event.set()
        self.resume_event.set()
        self._set_buttons("working")
        self.instr_var.set("")
        self.status_var.set("Пропускаю деталь…")

    def on_cancel(self) -> None:
        if not self._worker_alive():
            return
        if messagebox.askyesno(
                APP_TITLE,
                "Остановить выполнение?\nТекущая операция в КОМПАС завершится "
                "после текущего шага."):
            self.cancel_event.set()
            self.resume_event.set()   # если ждём на паузе — разбудить для отмены
            self._set_buttons("working")
            self.status_var.set("Останавливаю…")

    def _show_thickness_dialog(self, part_name: str, default: str) -> None:
        """
        Модальный диалог толщины (запрос рабочего потока): число мм ->
        («ok», значение); «Пропустить деталь» / «Остановить» — как кнопки.
        """
        win = tk.Toplevel(self)
        win.configure(bg=theme.BG)
        win.title("Толщина детали")
        win.geometry("400x190")
        win.resizable(False, False)
        win.grab_set()
        win.protocol("WM_DELETE_WINDOW", lambda: None)   # только по кнопкам

        ttk.Label(win, text=f"Деталь: {part_name}", wraplength=370,
                  justify="left",
                  font=theme.font(10, True)).pack(
                      anchor="w", padx=12, pady=(12, 6))
        ttk.Label(win,
                  text="Толщина материала, мм — DXF будет сохранён "
                       "в подпапку с этим номером:").pack(
                          anchor="w", padx=12)
        var = tk.StringVar(value=default or "")
        entry = ttk.Entry(win, textvariable=var)
        entry.pack(fill="x", padx=12, pady=(4, 8))
        entry.focus_set()
        entry.select_range(0, "end")

        def answer(result: Tuple[str, str]) -> None:
            if self.worker is not None:
                self.worker.thickness_result = result
                self.worker.thickness_event.set()
            if result[0] == "cancel":
                self.cancel_event.set()
                self.status_var.set("Останавливаю…")
            win.destroy()

        def on_ok() -> None:
            value = var.get().strip().replace(",", ".")
            if not re.fullmatch(r"\d+(\.\d+)?", value):
                messagebox.showwarning(
                    APP_TITLE,
                    "Введите толщину числом в миллиметрах,\n"
                    "например 4 или 4.5", parent=win)
                return
            answer(("ok", value))

        def on_skip() -> None:
            answer(("skip", ""))

        def on_stop() -> None:
            answer(("cancel", ""))

        entry.bind("<Return>", lambda _e: on_ok())
        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btns, text="ОК", command=on_ok).pack(side="left")
        ttk.Button(btns, text="Пропустить деталь",
                   command=on_skip).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="Остановить",
                   command=on_stop).pack(side="right")

    # ------------------------------------------------------------------
    # Обработка сообщений рабочего потока
    # ------------------------------------------------------------------

    def _poll_queue(self) -> None:
        """Единственный мост GUI <- рабочий поток (опрос каждые 100 мс)."""
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                self._handle_message(msg)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _handle_message(self, msg: Dict[str, Any]) -> None:
        mtype = msg.get("type")

        if mtype == "LOG":
            self._log_gui(msg.get("level", "info"), msg.get("text", ""))

        elif mtype == "STATUS":
            self.status_var.set(msg.get("text", ""))

        elif mtype == "PROGRESS":
            total = msg.get("total", 100)
            done = msg.get("done", 0)
            self.progress.configure(maximum=total, value=done)
            self.step_var.set(f"готово {done} из {total}" if total else "")

        elif mtype == "PART_START":
            self._mark_part_running(msg.get("uid", ""))

        elif mtype == "PAUSE":
            # Пауза: показать инструкцию, включить «Продолжить»/«Пропустить».
            self._set_buttons("paused")
            self.instr_var.set(
                f"{msg.get('title', '')}\n{msg.get('instruction', '')}\n\n"
                "Когда будете готовы, нажмите «Продолжить».")

        elif mtype == "ASK_THICKNESS":
            # Модальный запрос толщины детали (до её экспорта).
            self._show_thickness_dialog(msg.get("part", ""),
                                        msg.get("default", ""))

        elif mtype == "PART_UPDATE":
            self._update_part_row(msg.get("uid", ""),
                                  msg.get("status", "—"),
                                  msg.get("message", ""),
                                  msg.get("is_sheet"),
                                  msg.get("thickness", ""))
            self._update_license_label()

        elif mtype == "TRAVERSAL_DONE":
            parts = [PartInfo.from_dict(d) for d in msg.get("parts", [])]
            self._rebuild_part_list(parts)
            self._update_counts()
            self._set_buttons("idle")
            self.status_var.set("Анализ сборки завершён")
            if parts:
                messagebox.showinfo(APP_TITLE,
                                    f"Найдено деталей: {len(parts)}")
            else:
                messagebox.showwarning(
                    APP_TITLE,
                    "Детали в сборке не найдены (или обход не удался).\n"
                    "Добавьте файлы деталей кнопкой «Добавить файлы вручную».")

        elif mtype == "DONE":
            self._set_buttons("idle")
            self.instr_var.set("")
            self.status_var.set("Готово")
            self._update_license_label()
            self._log_gui("info", msg.get("text", ""))
            messagebox.showinfo(APP_TITLE, msg.get("text", ""))

        elif mtype == "FATAL":
            self._set_buttons("idle")
            self.instr_var.set("")
            self.status_var.set("Ошибка")
            self._log_gui("error", msg.get("text", ""))
            messagebox.showerror(APP_TITLE, msg.get("text", ""))

    # ------------------------------------------------------------------
    # Диалог лицензии и справка
    # ------------------------------------------------------------------

    def show_instructions(self) -> None:
        """Меню «Справка → Инструкция»: порядок работы."""
        messagebox.showinfo(f"Инструкция — {APP_TITLE}", STARTUP_INSTRUCTIONS)

    def show_about(self) -> None:
        """Меню «Справка → О программе»."""
        messagebox.showinfo(
            f"О программе — {APP_TITLE}",
            f"{APP_TITLE}\nВерсия {APP_VERSION}\n\n"
            "Пакетный экспорт деталей сборки КОМПАС-3D в DXF 1:1 (мм),\n"
            "включая исполнения и локальные детали (встроенные в сборку).\n\n"
            f"{VENDOR} — {SUPPORT_EMAIL}")

    def show_license_dialog(self) -> None:
        """Окно лицензии: статус, HWID для заказа, подключение license.key."""
        win = tk.Toplevel(self)
        win.configure(bg=theme.BG)
        win.title(f"Лицензия — {VENDOR}")
        win.geometry("560x400")
        win.resizable(False, False)
        win.grab_set()

        if self.lic.licensed:
            head = (f"ЛИЦЕНЗИЯ АКТИВНА\nКлиент: {self.lic.customer}\n"
                    f"Спасибо за покупку! ({VENDOR})")
        elif self.lic.corrupted:
            head = "Состояние пробной версии ПОВРЕЖДЕНО.\nЭкспорт заблокирован."
        else:
            head = (f"ПРОБНАЯ ВЕРСИЯ\nИспользовано экспортов: "
                    f"{self.lic.used} из {TRIAL_EXPORT_LIMIT}")

        ttk.Label(win, text=head, justify="left",
                  font=theme.font(10, True)).pack(anchor="w", padx=12,
                                                      pady=(12, 6))

        frm_hwid = ttk.Frame(win)
        frm_hwid.pack(fill="x", padx=12, pady=6)
        ttk.Label(frm_hwid, text="HWID этого компьютера:").pack(side="left")
        hwid_var = tk.StringVar(value=self.lic.get_hwid())
        entry = ttk.Entry(frm_hwid, textvariable=hwid_var, state="readonly")
        entry.pack(side="left", padx=8, fill="x", expand=True)

        def copy_hwid() -> None:
            self.clipboard_clear()
            self.clipboard_append(hwid_var.get())
            messagebox.showinfo(APP_TITLE, "HWID скопирован в буфер обмена.",
                                parent=win)

        ttk.Button(frm_hwid, text="Скопировать",
                   command=copy_hwid).pack(side="left")

        ttk.Label(win, justify="left", wraplength=530, text=(
            "Для получения лицензии:\n"
            f"  1. Скопируйте HWID и отправьте его вместе с именем заказчика "
            f"на {SUPPORT_EMAIL} ({VENDOR}).\n"
            "  2. Получите файл license.key и подключите его кнопкой ниже.\n"
            "Лицензия привязана к этому компьютеру (HWID) и не требует "
            "интернета."
        )).pack(anchor="w", padx=12, pady=6)

        def pick_license() -> None:
            path = filedialog.askopenfilename(
                title="Выберите файл лицензии",
                filetypes=[("Лицензия (*.key)", "*.key"),
                           ("Все файлы", "*.*")], parent=win)
            if not path:
                return
            ok, message = self.lic.activate_license(path)
            if ok:
                self.logger.info(f"Лицензия активирована: {self.lic.customer}")
                self._update_license_label()
                messagebox.showinfo(APP_TITLE, message, parent=win)
                win.destroy()
            else:
                messagebox.showerror(APP_TITLE, message, parent=win)

        ttk.Button(win, text="Выбрать файл лицензии…",
                   command=pick_license).pack(pady=8)
        ttk.Button(win, text="Закрыть", command=win.destroy).pack(pady=4)

    # ------------------------------------------------------------------
    # Закрытие приложения
    # ------------------------------------------------------------------

    def on_close(self) -> None:
        if self._worker_alive():
            if not messagebox.askyesno(
                    APP_TITLE,
                    "Выполняется фоновая операция.\n"
                    "Прервать и выйти? (КОМПАС завершит текущий шаг)") :
                return
            self.cancel_event.set()
            self.resume_event.set()
            self.worker.join(5000 / 1000.0)
            if self.worker.is_alive():
                messagebox.showwarning(
                    APP_TITLE,
                    "Рабочий поток ещё завершается — КОМПАС может остаться "
                    "занят на несколько секунд.")
        self.destroy()
