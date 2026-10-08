# -*- coding: utf-8 -*-
"""Оформление: палитра, шрифты и стили ttk в едином светлом стиле."""

import tkinter as tk
from tkinter import font as tkfont
from tkinter import ttk
from typing import Tuple

# Палитра.
ACCENT = "#1f6feb"
ACCENT_ACTIVE = "#1a5fd0"
ACCENT_DISABLED = "#b9c7df"
BG = "#f5f6f8"            # фон окна
CARD = "#ffffff"          # фон карточек
LINE = "#e3e6ea"          # рамки карточек и полей
ROW_ALT = "#fafbfc"       # чередование строк таблицы
TXT = "#1b1f24"
MUTED = "#6a737d"
OK = "#1a7f37"
OK_BG = "#e8f5ec"
WARN = "#9a6700"
WARN_BG = "#fff4d6"
ERR = "#b42318"
ERR_BG = "#fdecea"
SELECTED = "#dceaff"

# Цвет строки списка деталей по статусу.
STATUS_COLORS = {"OK": OK, "ERROR": ERR, "SKIP": WARN, "RUN": ACCENT,
                 "OFF": MUTED}


def _family(root: tk.Misc) -> str:
    """Segoe UI на Windows, иначе — первый доступный шрифт без засечек."""
    available = set(tkfont.families(root))
    for name in ("Segoe UI", "DejaVu Sans", "Helvetica", "Arial"):
        if name in available:
            return name
    return "TkDefaultFont"


FONT_FAMILY = "Segoe UI"


def font(size: int = 10, bold: bool = False) -> Tuple[str, int, str]:
    return (FONT_FAMILY, size, "bold" if bold else "normal")


def apply_theme(root: tk.Tk) -> ttk.Style:
    """Настроить стили ttk. Вызывается один раз при создании окна."""
    global FONT_FAMILY
    FONT_FAMILY = _family(root)

    root.configure(bg=BG)
    # Шрифт по умолчанию — через именованные шрифты Tk. option_add("*Font")
    # не подходит: он перебивает шрифты стилей ttk (жирные заголовки).
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                 "TkHeadingFont"):
        try:
            tkfont.nametofont(name).configure(family=FONT_FAMILY, size=10)
        except tk.TclError:
            pass

    style = ttk.Style(root)
    style.theme_use("clam")      # единственная тема, которую можно перекрасить
    style.configure(".", background=BG, foreground=TXT, font=font(10))

    style.configure("Card.TFrame", background=CARD)
    style.configure("Card.TLabel", background=CARD)
    style.configure("Title.TLabel", background=CARD, font=font(14, True))
    style.configure("Muted.TLabel", background=CARD, foreground=MUTED,
                    font=font(9))
    style.configure("Section.TLabel", background=CARD, font=font(10, True))
    style.configure("Instr.TLabel", background=CARD, foreground=ACCENT,
                    font=font(10))
    style.configure("Card.TCheckbutton", background=CARD)
    style.configure("Card.TRadiobutton", background=CARD)

    style.configure("TEntry", fieldbackground="#fbfcfd", bordercolor=LINE,
                    lightcolor=LINE, darkcolor=LINE, padding=5)
    style.map("TEntry", bordercolor=[("focus", ACCENT)],
              lightcolor=[("focus", ACCENT)])

    style.configure("TButton", background="#eef0f3", bordercolor=LINE,
                    lightcolor="#eef0f3", darkcolor="#eef0f3",
                    focuscolor="#eef0f3", padding=(12, 6))
    style.map("TButton",
              background=[("disabled", "#f3f4f6"), ("active", "#e2e5ea")],
              foreground=[("disabled", "#a5abb3")])
    style.configure("Ghost.TButton", background=CARD, lightcolor=CARD,
                    darkcolor=CARD, focuscolor=CARD)
    style.map("Ghost.TButton",
              background=[("disabled", CARD), ("active", "#f0f2f5")])
    style.configure("Accent.TButton", background=ACCENT, foreground="white",
                    bordercolor=ACCENT, lightcolor=ACCENT, darkcolor=ACCENT,
                    focuscolor=ACCENT, padding=(16, 7), font=font(10, True))
    style.map("Accent.TButton",
              background=[("disabled", ACCENT_DISABLED),
                          ("active", ACCENT_ACTIVE)],
              bordercolor=[("disabled", ACCENT_DISABLED)],
              lightcolor=[("disabled", ACCENT_DISABLED)],
              darkcolor=[("disabled", ACCENT_DISABLED)],
              foreground=[("disabled", "white")])

    style.configure("Thin.Horizontal.TProgressbar", background=ACCENT,
                    troughcolor="#e9edf2", bordercolor="#e9edf2",
                    lightcolor=ACCENT, darkcolor=ACCENT, thickness=6)

    style.configure("Parts.Treeview", background=CARD, fieldbackground=CARD,
                    foreground=TXT, rowheight=28, bordercolor=LINE,
                    font=font(10))
    style.configure("Parts.Treeview.Heading", background="#f0f2f5",
                    foreground=TXT, font=font(9, True), padding=(6, 6),
                    relief="flat", bordercolor=LINE)
    style.map("Parts.Treeview.Heading", background=[("active", "#e6e9ee")])
    style.map("Parts.Treeview", background=[("selected", SELECTED)],
              foreground=[("selected", TXT)])
    style.configure("Vertical.TScrollbar", background="#e9edf2",
                    troughcolor=CARD, bordercolor=CARD, arrowcolor=MUTED,
                    lightcolor="#e9edf2", darkcolor="#e9edf2")
    return style


def card(parent: tk.Misc, padding: int = 12) -> Tuple[tk.Frame, ttk.Frame]:
    """Карточка: белая панель с тонкой рамкой. -> (внешняя, внутренняя)."""
    outer = tk.Frame(parent, bg=LINE)
    inner = ttk.Frame(outer, style="Card.TFrame", padding=padding)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    return outer, inner
