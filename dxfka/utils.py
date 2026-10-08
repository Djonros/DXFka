# -*- coding: utf-8 -*-
"""Утилиты общего назначения, журнал, настройки, COM-коллекции."""

import json
import logging
import logging.handlers
import os
import re
from typing import Any, Dict, List, Optional

from dxfka.config import (
    APP_DIR, LOG_FILE, MAX_NAME_LEN, SETTINGS_FILE, SKIP_PATH_MARKERS,
    STATE_DIR)

# ======================================================================
# УТИЛИТЫ ОБЩЕГО НАЗНАЧЕНИЯ
# ======================================================================

def app_dir() -> str:
    """Каталог приложения: для exe — папка с exe, иначе корень репозитория."""
    return APP_DIR


def sanitize_filename(name: str) -> str:
    """Заменяет запрещённые в Windows имена символы, обрезает длину."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(".")
    return cleaned[:MAX_NAME_LEN].strip()


def is_standard_part_path(path: str) -> bool:
    """True, если путь похож на стандартное/библиотечное изделие (пропускаем)."""
    low = os.path.normcase(path)
    return any(marker in low for marker in SKIP_PATH_MARKERS)


def safe_str(obj: Any, attr: str, default: str = "") -> str:
    """Безопасное чтение строкового свойства COM-объекта (может отсутствовать)."""
    try:
        value = getattr(obj, attr)
        if callable(value):          # редко: свойство может прийти как метод
            value = value()
        if value is None:
            return default
        return str(value)
    except Exception:
        return default


def safe_int(obj: Any, attr: str, default: int = 0) -> int:
    """Безопасное чтение целочисленного свойства COM-объекта."""
    try:
        value = getattr(obj, attr)
        if callable(value):
            value = value()
        return int(value)
    except Exception:
        return default


# ----------------------------------------------------------------------
# Универсальные помощники для перебора COM-коллекций КОМПАС.
# Разные версии API дают доступ по разным индексаторам и с разной базой
# (0 или 1), поэтому пробуем несколько вариантов.
# TODO_KOMPAS: точная схема индексации каждой коллекции не документирована;
# при необходимости уточнить по SDK конкретной версии КОМПАС.
# ----------------------------------------------------------------------

def com_count(collection: Any) -> Optional[int]:
    """Количество элементов коллекции (Count/count/GetCount)."""
    if collection is None:
        return None
    for attr in ("Count", "count"):
        try:
            value = getattr(collection, attr)
            if value is not None:
                return int(value)
        except Exception:
            continue
    try:
        return int(collection.GetCount())
    except Exception:
        return None


def com_item(collection: Any, index: int) -> Any:
    """Элемент коллекции по индексу: [] / Item / Part / Component / GetItem."""
    try:
        return collection[index]
    except Exception:
        pass
    for method in ("Item", "Part", "Component", "GetItem"):
        try:
            return getattr(collection, method)(index)
        except Exception:
            continue
    return None


def com_items(collection: Any) -> list:
    """
    Список всех элементов коллекции. Пробуем 0-базу и 1-базу:
    если первый индекс не даёт элемент — переключаемся на вторую базу.
    """
    count = com_count(collection)
    if not count:
        return []
    for base in (0, 1):
        items = []
        for i in range(base, base + count):
            item = com_item(collection, i)
            if item is None:
                items = []
                break
            items.append(item)
        if items:
            return items
    return []


def setup_logging() -> logging.Logger:
    """Журнал в файл рядом с приложением (ротация 2 МБ x 3)."""
    logger = logging.getLogger("kompas_dxf_exporter")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        try:
            handler = logging.handlers.RotatingFileHandler(
                LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
            handler.setFormatter(
                logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s"))
            logger.addHandler(handler)
        except Exception:
            pass  # без файла журнал всё равно идёт в GUI
    return logger


def load_settings() -> Dict[str, Any]:
    """Настройки GUI из %APPDATA% (шаблон фрагмента, папка вывода)."""
    try:
        with open(os.path.join(STATE_DIR, SETTINGS_FILE),
                  encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data: Dict[str, Any]) -> None:
    try:
        os.makedirs(STATE_DIR, exist_ok=True)
        with open(os.path.join(STATE_DIR, SETTINGS_FILE), "w",
                  encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
    except Exception:
        pass  # настройки не критичны


# ----------------------------------------------------------------------
# Толщина по материалу и размер DXF
# ----------------------------------------------------------------------

# «Лист$d8 ГОСТ 19903-2015;…», «Лист 4 ГОСТ 8568-77», «Лист Б-ПН-О-10 …»
# ($d — так КОМПАС пишет перенос строки в наименовании материала).
_SHEET_MATERIAL_RE = re.compile(
    r"^\s*лист\b[^0-9;]{0,16}?(\d+(?:[.,]\d+)?)", re.IGNORECASE)


def thickness_from_material(material: str) -> str:
    """Толщина листа из строки материала («8», «1.5»); "" — не лист."""
    match = _SHEET_MATERIAL_RE.match(material or "")
    if not match:
        return ""
    value = float(match.group(1).replace(",", "."))
    if value <= 0:
        return ""
    return f"{value:.3f}".rstrip("0").rstrip(".")


def dxf_extent_area(path: str) -> float:
    """
    Площадь габарита геометрии DXF (по координатам 10/20 и 11/21
    в секции ENTITIES); 0 — файл пуст/не читается.
    """
    try:
        with open(path, "r", encoding="latin-1") as fh:
            lines = [line.strip() for line in fh]
    except OSError:
        return 0.0
    xs: List[float] = []
    ys: List[float] = []
    in_entities = False
    for i in range(0, len(lines) - 1, 2):
        code, value = lines[i], lines[i + 1]
        if code == "2" and value == "ENTITIES":
            in_entities = True
        elif code == "0" and value == "ENDSEC":
            in_entities = False
        elif in_entities and code in ("10", "11", "20", "21"):
            try:
                (xs if code in ("10", "11") else ys).append(float(value))
            except ValueError:
                pass
    if not xs or not ys:
        return 0.0
    return (max(xs) - min(xs)) * (max(ys) - min(ys))
