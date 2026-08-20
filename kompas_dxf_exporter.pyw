# -*- coding: utf-8 -*-
# Запуск kompas_dxf_exporter БЕЗ консольного окна (двойной клик).
# Рядом с этим файлом должен лежать kompas_dxf_exporter.py
# (либо установленный Python 3.10+ с пакетом pywin32).
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import kompas_dxf_exporter
    kompas_dxf_exporter.main()
except Exception:
    # Любая ошибка старта — аккуратное окно вместо молчаливого закрытия.
    try:
        import tkinter.messagebox as mb
        mb.showerror("kompas_dxf_exporter — Djonros",
                     "Не удалось запустить приложение:\n\n"
                     + traceback.format_exc())
    except Exception:
        raise
