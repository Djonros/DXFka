# -*- coding: utf-8 -*-
"""
DXFka — пакетный экспорт DXF из деталей сборки КОМПАС-3D v24.
Автор и поставщик: Djonros (djonros@gmail.com).

Возможности:
  * выбор файла сборки (.a3d) и рекурсивный обход всех входящих деталей (.m3d),
    включая вложенные подсборки, с дедупликацией и пропуском стандартных изделий;
  * определение листовых деталей через API листового моделирования
    (ISheetMetalBody) и автоматическое построение развертки (Straighten);
  * пауза на каждой детали: пользователь проверяет развертку / поворачивает
    модель (опционально сохранив вид с именем «DXF_Export»), затем нажимает
    «Продолжить» в окне программы;
  * создание чертежа-фрагмента 1:1 БЕЗ рамки и штампа и экспорт в DXF
    (ksSaveToDXF) — только геометрия, единицы мм;
  * подробный журнал (kompas_dxf_exporter.log) и сводка по результатам;
  * коммерциализация: пробная версия с лимитом экспортов, оффлайн-лицензия
    (файл license.key), привязанная к компьютеру (HWID).

ЗАПУСК (Python 3.10+, Windows 10/11):
  1) Установить зависимости:      pip install pywin32
  2) Запустить с консолью:        python kompas_dxf_exporter.py
     или без консоли (двойной клик): kompas_dxf_exporter.pyw
     или готовым исполняемым файлом: kompas_dxf_exporter.exe (PyInstaller)
  3) Должен быть установлен КОМПАС-3D v24. Программа подключается к уже
     запущенному экземпляру КОМПАС (ProgID KOMPAS.Application.7/8), иначе
     запускает новый.
  ВАЖНО: если КОМПАС и скрипт запущены с разными правами (админ/пользователь),
  подключение может не сработать — запускайте их с одинаковыми правами.

ЛИЦЕНЗИРОВАНИЕ:
  * Пробная версия: TRIAL_EXPORT_LIMIT успешных экспортов, затем требуется
    лицензия. Счётчик хранится в %APPDATA%\\KompasDxfExporter (файл + реестр).
  * Для покупки лицензии: откройте в программе меню «Лицензия…», скопируйте
    HWID и отправьте на djonros@gmail.com (Djonros). Полученный файл
    license.key подключите в том же диалоге.
"""

# ======================================================================
# ИМПОРТЫ СТАНДАРТНОЙ БИБЛИОТЕКИ
# ======================================================================
import os
import sys
import re
import json
import time
import hmac
import uuid
import queue
import shutil
import hashlib
import logging
import logging.handlers
import tempfile
import traceback
import threading
import winreg
from dataclasses import dataclass, field
from typing import Optional, List, Tuple, Callable, Dict, Any

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# --- pywin32: нужен для работы с COM (КОМПАС). Может отсутствовать. ---
PYWIN32_OK = True
try:
    import pythoncom          # инициализация COM в потоке
    import pywintypes         # типы исключений COM
    from win32com.client import Dispatch, GetActiveObject, gencache, CastTo
except ImportError:  # pragma: no cover - на машине без pywin32
    PYWIN32_OK = False

# ======================================================================
# КОНСТАНТЫ: приложение, брендинг
# ======================================================================

APP_TITLE = "DXFka — Экспорт DXF из КОМПАС-3D"
VENDOR = "Djonros"
SUPPORT_EMAIL = "djonros@gmail.com"
LICENSE_CONTACT_HINT = (f"Для получения лицензии обратитесь: {VENDOR} — {SUPPORT_EMAIL}")
TRIAL_EXHAUSTED_MSG = ("Лимит пробной версии исчерпан.\n" + LICENSE_CONTACT_HINT)
TRIAL_CORRUPTED_MSG = ("Состояние пробной версии повреждено "
                       "(обнаружена попытка сброса счётчика).\nЭкспорт заблокирован. "
                       + LICENSE_CONTACT_HINT)

# ======================================================================
# КОНСТАНТЫ: лицензирование / пробная версия
# ======================================================================

TRIAL_EXPORT_LIMIT = 10                     # сколько успешных экспортов даёт пробная версия
PRODUCT_ID = "kompas_dxf_exporter"          # идентификатор продукта в license.key
STATE_DIR = os.path.join(os.environ.get("APPDATA", tempfile.gettempdir()),
                         "KompasDxfExporter")            # %APPDATA%\KompasDxfExporter
STATE_FILE = "state.json"                   # счётчик триала (файл)
SETTINGS_FILE = "settings.json"             # настройки GUI (шаблон, папка вывода)
LICENSE_FILE = "license.key"                # файл лицензии
REG_KEY = r"Software\KompasDxfExporter"     # зеркало счётчика в реестре (HKCU)

# Секрет для подписи лицензий и контрольных сумм состояния.
# Живёт в license_secret.py (в .gitignore, в репозиторий не входит;
# см. license_secret.example.py). Без него — режим разработки:
# триал работает, лицензии не проверяются.
# TODO_SECURITY: для серьёзной коммерческой защиты дополнительно применить
# PyArmor/похожие обфускаторы и не полагаться только на HMAC.
try:
    from license_secret import SECRET
except ImportError:
    SECRET = b"DXFka-dev-fallback-secret"

# ======================================================================
# КОНСТАНТЫ: КОМПАС API и экспорт
# ======================================================================

# ProgID API7 в порядке приоритета (требование: сначала .7, затем .8).
PROGIDS_API7 = ["KOMPAS.Application.7", "KOMPAS.Application.8"]
PROGID_API5 = "Kompas.Application.5"        # API5 (KompasObject) — 2D/чертежи

# GUID библиотек типов КОМПАС (проверено на v24):
TYPELIB_CONSTANTS = "{75C9F5D0-B5B8-4526-8681-9903C567D2ED}"  # константы ks*
TYPELIB_API5 = "{0422828C-F174-495E-AC5D-D31014DBBE87}"       # API5 (KompasObject)
TYPELIB_API7 = "{69AC2981-37C0-4379-84FD-5DD2F3C0A520}"       # API7 (IApplication)

# Локальные значения констант — используются, если загрузка typelib не удалась
# (актуально для замороженного exe, где gencache может не работать).
KO_DOCUMENT_PARAM = 35                       # ko_DocumentParam
KO_ASSOCIATION_VIEW_PARAM = 122              # ko_AssociationViewParam

DOC_TYPE_FRAGMENT = 2                        # ksDocumentFragment: фрагмент БЕЗ рамки и штампа
DOC_TYPE_DRAWING = 1                         # ksDocumentDrawing (не используется, для справки)

FLAT_PROJECTION_NAME = "#Развертка"          # имя проекции-развертки (проверено работает)
FRONT_PROJECTION_NAME = "#Спереди"           # стандартная проекция (fallback для нелистовых)
SAVED_VIEW_NAME = "DXF_Export"               # имя сохранённого пользователем вида в 3D

# Пути, содержащие эти маркеры, считаем стандартными/библиотечными изделиями
# (Ascon/Program Files) и пропускаем при обходе сборки.
SKIP_PATH_MARKERS = ("program files", "ascon")

# Тайминги и повторы (КОМПАС иногда «задумывается» на больших моделях):
OPEN_RETRY_COUNT = 20          # попыток открыть документ
OPEN_RETRY_WAIT_S = 3.0        # пауза между попытками, сек (итого до ~60 с)
REBUILD_WAIT_S = 3.0           # ожидание перестройки модели после Straighten
AFTER_CREATE_DOC_S = 0.5       # пауза после ksCreateDocument
AFTER_VIEW_S = 1.0             # пауза после создания вида перед сохранением DXF
MAX_ASSEMBLY_DEPTH = 10        # защита от циклов в дереве подсборок
MAX_NAME_LEN = 150             # максимальная длина имени файла DXF (без .dxf)

LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(sys.executable
                         if getattr(sys, "frozen", False) else __file__)),
                        "kompas_dxf_exporter.log")
DEFAULT_OUT_SUBDIR = "DXF_Out"

STARTUP_INSTRUCTIONS = (
    "Порядок работы:\n"
    "  1. Укажите файл сборки (.a3d) и папку для DXF.\n"
    "  2. Нажмите «Загрузить детали» — программа проанализирует сборку. "
    "Если детали не найдены — добавьте их кнопкой «Добавить файлы вручную».\n"
    "  3. Отметьте нужные детали и нажмите «Начать экспорт».\n"
    "  4. Для КАЖДОЙ детали программа остановится: листовая — проверьте "
    "развертку; нелистовая — поверните модель (видно в КОМПАС), при желании "
    "сохраните вид с именем «DXF_Export». Затем нажмите «Продолжить».\n"
    "  5. Результат: DXF 1:1 в мм, только геометрия (фрагмент без рамки).\n"
    f"  Поддержка и лицензии: {VENDOR} — {SUPPORT_EMAIL}"
)


# ======================================================================
# УТИЛИТЫ ОБЩЕГО НАЗНАЧЕНИЯ
# ======================================================================

def app_dir() -> str:
    """Каталог приложения: для exe — папка с exe, иначе папка скрипта."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


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


# ======================================================================
# ЛИЦЕНЗИРОВАНИЕ: HWID, состояние триала, файл лицензии
# ======================================================================

class LicenseManager:
    """
    Управление пробной версией и лицензией.

    Пробная версия: TRIAL_EXPORT_LIMIT успешных экспортов. Счётчик хранится
    в двух местах (state.json + реестр HKCU), каждое значение защищено
    HMAC-подписью. Берётся максимум из корректных значений; если оба счётчика
    исчезли/повреждены при наличии маркера установки — состояние считается
    повреждённым и экспорт блокируется (fail-closed) до ввода лицензии.

    Лицензия: JSON-файл license.key с полями product, customer, hwid, issued
    и подписью sig = HMAC-SHA256(SECRET, "product|customer|hwid|issued").
    Действительна только на компьютере с совпадающим HWID.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self._hwid: Optional[str] = None
        self.licensed: bool = False          # активна ли лицензия
        self.customer: str = ""              # имя клиента из license.key
        self.used: int = 0                   # использовано экспортов триала
        self.corrupted: bool = False         # повреждено состояние триала
        self._reserved: int = 0              # экспорты «в полёте» (резерв слотов)

    # ------------------------- HWID (привязка к ПК) -------------------------

    def get_hwid(self) -> str:
        """Идентификатор компьютера XXXX-XXXX-XXXX-XXXX (кэшируется)."""
        if self._hwid is None:
            self._hwid = self._compute_hwid()
        return self._hwid

    @staticmethod
    def _machine_guid() -> str:
        """MachineGuid из реестра (учитываем 64-битный вид реестра)."""
        for access in (winreg.KEY_READ | winreg.KEY_WOW64_64KEY, winreg.KEY_READ):
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Cryptography", 0, access)
                try:
                    value, _ = winreg.QueryValueEx(key, "MachineGuid")
                    return str(value)
                finally:
                    winreg.CloseKey(key)
            except OSError:
                continue
        return ""

    @staticmethod
    def _volume_serial() -> str:
        """Серийный номер тома C: через pywin32 (если доступен)."""
        try:
            import win32api
            return str(win32api.GetVolumeInformation("C:\\")[0])
        except Exception:
            return ""

    @classmethod
    def _compute_hwid(cls) -> str:
        """
        HWID = первые 16 hex-символов SHA256(MachineGuid + VolumeSerial C:).
        При полном отказе обоих источников — MAC-адрес (uuid.getnode).
        """
        guid = cls._machine_guid()
        vol = cls._volume_serial()
        raw = f"{guid}|{vol}".encode("utf-8")
        if not guid and not vol:
            raw = str(uuid.getnode()).encode("utf-8")
        digest = hashlib.sha256(raw).hexdigest()[:16].upper()
        return "-".join(digest[i:i + 4] for i in range(0, 16, 4))

    # ------------------------- подписи (HMAC) -------------------------

    @staticmethod
    def license_sig(product: str, customer: str, hwid: str, issued: str) -> str:
        """Подпись лицензии. Тот же алгоритм используется в keygen.py."""
        payload = f"{product}|{customer}|{hwid}|{issued}".encode("utf-8")
        return hmac.new(SECRET, payload, hashlib.sha256).hexdigest()

    @staticmethod
    def _trial_guard(hwid: str, used: int) -> str:
        """Контрольная сумма счётчика триала (защита от подмены)."""
        payload = f"{hwid}|{used}|trial".encode("utf-8")
        return hmac.new(SECRET, payload, hashlib.sha256).hexdigest()

    # ------------------------- загрузка состояния -------------------------

    def load(self) -> None:
        """Полная загрузка: HWID → состояние триала → файл лицензии."""
        with self._lock:
            self._hwid = self._compute_hwid()
            self._load_trial_state()
            self._load_license_file()

    def _load_trial_state(self) -> None:
        """Чтение счётчика из state.json и реестра, защита от сброса."""
        file_used, file_bad = self._read_state_file()
        reg_used, reg_bad = self._read_state_registry()
        marker_ok = self._read_install_marker()

        valid = [u for u in (file_used, reg_used) if u is not None]
        any_bad = file_bad or reg_bad

        if valid:
            # Берём максимум корректных значений и синхронизируем хранилища.
            self.used = max(valid)
            self._write_state()
        elif any_bad:
            # Что-то есть, но всё повреждено — подозрение на подмену.
            self.corrupted = True
        elif marker_ok:
            # Программа уже запускалась (маркер установлен), но оба счётчика
            # исчезли — типичный признак ручного сброса. Fail-closed.
            self.corrupted = True
        else:
            # Чистая первая установка.
            self.used = 0
            self._write_state()
            self._write_install_marker()

    def _read_state_file(self) -> Tuple[Optional[int], bool]:
        """-> (использовано | None, повреждён ли файл)."""
        path = os.path.join(STATE_DIR, STATE_FILE)
        if not os.path.isfile(path):
            return None, False
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            used = int(data["used"])
            guard = str(data["guard"])
            if hmac.compare_digest(guard, self._trial_guard(self.get_hwid(), used)):
                return used, False
            return None, True
        except Exception:
            return None, True

    def _write_state(self) -> None:
        """Сохранить счётчик в файл и реестр (оба места сразу)."""
        used = self.used
        guard = self._trial_guard(self.get_hwid(), used)
        try:
            os.makedirs(STATE_DIR, exist_ok=True)
            tmp = os.path.join(STATE_DIR, STATE_FILE + ".tmp")
            with open(tmp, "w", encoding="utf-8") as fh:
                json.dump({"used": used, "guard": guard}, fh)
            os.replace(tmp, os.path.join(STATE_DIR, STATE_FILE))
        except Exception:
            pass  # файл может быть недоступен — останется реестр
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
                winreg.SetValueEx(key, "ts", 0, winreg.REG_SZ, f"{used}|{guard}")
        except OSError:
            pass

    def _read_state_registry(self) -> Tuple[Optional[int], bool]:
        """Счётчик-зеркало из реестра: значение 'ts' = 'used|guard'."""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
                raw, _ = winreg.QueryValueEx(key, "ts")
            used_str, guard = str(raw).split("|", 1)
            used = int(used_str)
            if hmac.compare_digest(guard, self._trial_guard(self.get_hwid(), used)):
                return used, False
            return None, True
        except FileNotFoundError:
            return None, False
        except OSError:
            return None, False
        except Exception:
            return None, True

    def _write_install_marker(self) -> None:
        """Маркер 'программа уже запускалась' (для обнаружения сброса счётчика)."""
        payload = f"{self.get_hwid()}|installed".encode("utf-8")
        marker = hmac.new(SECRET, payload, hashlib.sha256).hexdigest()
        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
                winreg.SetValueEx(key, "in", 0, winreg.REG_SZ, marker)
        except OSError:
            pass

    def _read_install_marker(self) -> bool:
        payload = f"{self.get_hwid()}|installed".encode("utf-8")
        expected = hmac.new(SECRET, payload, hashlib.sha256).hexdigest()
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_KEY) as key:
                raw, _ = winreg.QueryValueEx(key, "in")
            return hmac.compare_digest(str(raw), expected)
        except Exception:
            return False

    # ------------------------- лицензия -------------------------

    def _validate_license_data(self, data: Dict[str, Any]) -> Tuple[bool, str]:
        """Проверка полей и подписи license-данных. hwid сверяется с текущим."""
        try:
            product = str(data.get("product", ""))
            customer = str(data.get("customer", "")).strip()
            hwid = str(data.get("hwid", "")).strip().upper()
            issued = str(data.get("issued", ""))
            sig = str(data.get("sig", ""))
        except Exception:
            return False, "Файл лицензии имеет неверный формат."
        if not customer or not hwid or not issued or not sig:
            return False, "В файле лицензии не хватает полей."
        if product != PRODUCT_ID:
            return False, "Ключ выдан для другого продукта."
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", issued):
            return False, "Некорректная дата выдачи в ключе."
        expected = self.license_sig(product, customer, hwid, issued)
        if not hmac.compare_digest(sig, expected):
            return False, "Подпись ключа недействительна (файл повреждён)."
        if hwid != self.get_hwid():
            return False, ("Ключ выдан для другого компьютера "
                           f"(HWID ключа: {hwid}, этот ПК: {self.get_hwid()}).")
        return True, customer

    def _load_license_file(self) -> None:
        """Поиск license.key: рядом с приложением, затем в %APPDATA%."""
        for candidate in (os.path.join(app_dir(), LICENSE_FILE),
                          os.path.join(STATE_DIR, LICENSE_FILE)):
            if not os.path.isfile(candidate):
                continue
            try:
                with open(candidate, "r", encoding="utf-8") as fh:
                    data = json.load(fh)
            except Exception:
                continue
            ok, info = self._validate_license_data(data)
            if ok:
                self.licensed = True
                self.customer = info
                return

    def activate_license(self, path: str) -> Tuple[bool, str]:
        """
        Подключение лицензии из указанного файла.
        При успехе копирует ключ в %APPDATA% (переживает обновления exe).
        """
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
        except Exception as exc:
            return False, f"Не удалось прочитать файл лицензии: {exc}"
        ok, info = self._validate_license_data(data)
        if not ok:
            return False, info
        with self._lock:
            self.licensed = True
            self.customer = info
            try:
                os.makedirs(STATE_DIR, exist_ok=True)
                shutil.copyfile(path, os.path.join(STATE_DIR, LICENSE_FILE))
            except Exception:
                pass  # ключ рядом с exe тоже подхватится при следующем запуске
        return True, (f"Лицензия активирована. Клиент: {info}. "
                      f"Спасибо за покупку! ({VENDOR})")

    # ------------------------- учёт экспортов -------------------------

    @property
    def remaining(self) -> int:
        """Сколько успешных экспортов осталось в пробной версии."""
        with self._lock:
            if self.licensed:
                return TRIAL_EXPORT_LIMIT  # формально «не ограничено»
            return max(0, TRIAL_EXPORT_LIMIT - self.used - self._reserved)

    def reserve(self) -> bool:
        """
        Зарезервировать слот под один экспорт (вызывается воркером ДО детали).
        False — экспорт запрещён (лимит исчерпан/повреждено/нет лицензии смысла).
        """
        with self._lock:
            if self.licensed:
                return True
            if self.corrupted:
                return False
            if self.used + self._reserved < TRIAL_EXPORT_LIMIT:
                self._reserved += 1
                return True
            return False

    def commit(self) -> None:
        """Успешный экспорт: снять резерв, увеличить счётчик (только в триале)."""
        with self._lock:
            self._reserved = max(0, self._reserved - 1)
            if not self.licensed:
                self.used += 1
                self._write_state()

    def release(self) -> bool:
        """Экспорт не состоялся (ошибка/пропуск/отмена): вернуть резерв."""
        with self._lock:
            self._reserved = max(0, self._reserved - 1)
            return True

    def status_text(self) -> str:
        """Текст для статус-строки GUI."""
        with self._lock:
            if self.licensed:
                return f"Лицензия: {self.customer} — {VENDOR}"
            if self.corrupted:
                return ("Состояние пробной версии повреждено — нужна лицензия "
                        f"({SUPPORT_EMAIL})")
            left = max(0, TRIAL_EXPORT_LIMIT - self.used - self._reserved)
            return (f"Пробная версия: осталось {left} из "
                    f"{TRIAL_EXPORT_LIMIT} экспортов")


# ======================================================================
# МОДЕЛЬ ДАННЫХ ДЕТАЛИ
# ======================================================================

@dataclass
class PartInfo:
    """Сведения об одной детали для списка в GUI."""
    file_path: str
    name: str = ""                # наименование из сборки
    marking: str = ""             # обозначение из сборки
    material: str = ""            # материал
    is_sheet: Optional[bool] = None   # None — неизвестно (уточнится при экспорте)
    status: str = "PENDING"       # PENDING | OK | ERROR | SKIP
    message: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "file_path": self.file_path,
            "name": self.name,
            "marking": self.marking,
            "material": self.material,
            "is_sheet": self.is_sheet,
            "status": self.status,
            "message": self.message,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PartInfo":
        return cls(
            file_path=data.get("file_path", ""),
            name=data.get("name", ""),
            marking=data.get("marking", ""),
            material=data.get("material", ""),
            is_sheet=data.get("is_sheet"),
        )


def build_dxf_basename(part: PartInfo) -> str:
    """Базовое имя DXF: «обозначение наименование», иначе имя исходного файла."""
    base = f"{part.marking} {part.name}".strip()
    if not base:
        base = os.path.splitext(os.path.basename(part.file_path))[0]
    return sanitize_filename(base) or "part"


def part_type_text(is_sheet: Optional[bool]) -> str:
    """Текст колонки «тип» (звёздочка = тип определён по материалу, не по API)."""
    if is_sheet is None:
        return "?"
    return "листовая" if is_sheet else "нелистовая"


# ======================================================================
# ИСКЛЮЧЕНИЯ УПРАВЛЕНИЯ ПОТОКОМ ЭКСПОРТА
# ======================================================================

class KompasConnectionError(ConnectionError):
    """Не удалось подключиться к КОМПАС (не установлен / нет pywin32 / занят)."""


class CancelledByUser(Exception):
    """Пользователь нажал «Остановить» во время паузы/работы."""


class SkippedByUser(Exception):
    """Пользователь нажал «Пропустить деталь» на паузе."""


# ======================================================================
# ПОДКЛЮЧЕНИЕ К КОМПАС (выполняется ТОЛЬКО в рабочем потоке)
# ======================================================================

class KompasConnection:
    """
    Подключение к КОМПАС-3D: API5 (KompasObject, 2D-документы/чертежи)
    и API7 (IApplication, 3D-модели/сборки).

    Проверенная схема (batch_export_dxf_v3.py): раннее связывание через
    gencache.EnsureModule(GUID, 0, 1, 0) + QueryInterface к нужному
    интерфейсу. GetActiveObject подключается к работающему КОМПАС,
    Dispatch запускает новый экземпляр.
    """

    def __init__(self, log: Callable[[str, str], None]):
        self.log = log
        self.constants = None        # модуль констант ks*/ko*
        self.api5_module = None      # ранние обёртки API5
        self.api7_module = None      # ранние обёртки API7
        self.kompas_object = None    # API5: KompasObject
        self.application = None      # API7: IApplication
        self.started_by_us = False   # запустили ли КОМПАС сами
        # Числовые константы (fallback, если typelib не загрузилась):
        self.KO_DOCUMENT_PARAM = KO_DOCUMENT_PARAM
        self.KO_ASSOCIATION_VIEW_PARAM = KO_ASSOCIATION_VIEW_PARAM

    # ------------------------------------------------------------------

    def connect(self) -> None:
        """Подключение/запуск КОМПАС. Бросает KompasConnectionError."""
        if not PYWIN32_OK:
            raise KompasConnectionError(
                "Не установлен пакет pywin32. Выполните: pip install pywin32")

        # Инициализация COM для текущего потока (рабочий поток GUI).
        pythoncom.CoInitialize()

        # 1) Библиотеки типов. При отказе (например, в замороженном exe)
        #    продолжаем на динамическом диспатче с локальными константами.
        try:
            self.constants = gencache.EnsureModule(TYPELIB_CONSTANTS, 0, 1, 0).constants
        except Exception as exc:
            self.log("warning", f"Не загрузилась typelib констант ({exc}); "
                                "использую локальные значения констант")
        try:
            self.api5_module = gencache.EnsureModule(TYPELIB_API5, 0, 1, 0)
        except Exception as exc:
            self.log("warning", f"Не загрузилась typelib API5 ({exc}); "
                                "2D-структуры будут динамическими")
        try:
            self.api7_module = gencache.EnsureModule(TYPELIB_API7, 0, 1, 0)
        except Exception as exc:
            self.log("warning", f"Не загрузилась typelib API7 ({exc}); "
                                "API7 будет динамическим")

        if self.constants is not None:
            self.KO_DOCUMENT_PARAM = int(getattr(
                self.constants, "ko_DocumentParam", KO_DOCUMENT_PARAM))
            self.KO_ASSOCIATION_VIEW_PARAM = int(getattr(
                self.constants, "ko_AssociationViewParam", KO_ASSOCIATION_VIEW_PARAM))

        # 2) API7: сначала подключаемся к работающему экземпляру.
        raw_app = None
        last_error = ""
        for progid in PROGIDS_API7:
            try:
                raw_app = GetActiveObject(progid)
                self.log("info", f"Подключено к работающему КОМПАС ({progid})")
                break
            except Exception as exc:
                last_error = str(exc)
        # Не нашли запущенный — создаём новый экземпляр.
        if raw_app is None:
            for progid in PROGIDS_API7:
                try:
                    raw_app = Dispatch(progid)
                    self.started_by_us = True
                    self.log("info", f"Запущен новый экземпляр КОМПАС ({progid})")
                    break
                except Exception as exc:
                    last_error = str(exc)
        if raw_app is None:
            raise KompasConnectionError(
                "КОМПАС-3D не найден (KOMPAS.Application.7/8). "
                f"Убедитесь, что КОМПАС v24 установлен. Последняя ошибка: {last_error}")

        # 3) Приводим к раннему интерфейсу IApplication (API7).
        try:
            self.application = self.api7_module.IApplication(
                raw_app._oleobj_.QueryInterface(
                    self.api7_module.IApplication.CLSID, pythoncom.IID_IDispatch))
        except Exception:
            # Fallback: динамический диспатч (методы те же, без ранних типов).
            self.application = raw_app
            self.log("warning", "IApplication: использую динамический диспатч")

        # 4) API5 (KompasObject) — тот же процесс, что и API7.
        try:
            raw5 = Dispatch(PROGID_API5)
            self.kompas_object = self.api5_module.KompasObject(
                raw5._oleobj_.QueryInterface(
                    self.api5_module.KompasObject.CLSID, pythoncom.IID_IDispatch))
        except Exception as exc:
            raise KompasConnectionError(
                f"Не удалось подключить API5 ({PROGID_API5}): {exc}")

        # 5) Окно КОМПАС должно быть видно: пользователь работает с моделью
        #    на паузах. Свойство может игнорироваться — не критично.
        try:
            self.application.Visible = True
        except Exception:
            self.log("warning", "Не удалось сделать окно КОМПАС видимым")

    # ------------------------------------------------------------------

    def pump(self, seconds: float) -> None:
        """Прокачка сообщений COM + ожидание (КОМПАС «дышит», GUI не вешаем)."""
        end = time.time() + seconds
        while time.time() < end:
            try:
                pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(0.2)


# ======================================================================
# ОБХОД СБОРКИ: получение списка деталей
# ======================================================================

class AssemblyWalker:
    """
    Рекурсивный обход сборки .a3d с извлечением всех уникальных деталей .m3d.

    Проверенный путь kompas_sheet_export.py (TopPart.PartsEx) на реальной
    сборке v24 вернул пустую коллекцию, поэтому перебираем несколько
    стратегий доступа к компонентам и логируем результат каждой.
    TODO_KOMPAS: уточнить по документации КОМПАС v24 корректный способ
    перечисления компонентов (IParts7.PartsEx / Parts / Components).
    """

    def __init__(self, conn: KompasConnection,
                 log: Callable[[str, str], None],
                 cancel_event: Optional[threading.Event] = None):
        self.conn = conn
        self.log = log
        self.cancel_event = cancel_event

    # ------------------------------------------------------------------

    def _cancelled(self) -> bool:
        return bool(self.cancel_event and self.cancel_event.is_set())

    def walk(self, assembly_path: str) -> List[PartInfo]:
        """Точка входа: возвращает список уникальных деталей сборки."""
        parts: List[PartInfo] = []
        seen = set()
        self._walk_file(assembly_path, parts, seen, depth=0)
        self.log("info", f"Обход завершён. Найдено уникальных деталей: {len(parts)}")
        return parts

    # ------------------------------------------------------------------

    def _open_document(self, path: str) -> Any:
        """Открытие документа с повторами (КОМПАС может быть занят)."""
        for attempt in range(1, OPEN_RETRY_COUNT + 1):
            if self._cancelled():
                return None
            try:
                doc = self.conn.application.Documents.Open(path)
                if doc is not None:
                    time.sleep(1.0)
                    return doc
            except Exception as exc:
                self.log("warning",
                         f"Открытие не удалось (попытка {attempt}): {exc}")
            self.conn.pump(OPEN_RETRY_WAIT_S)
        return None

    def _close_document(self, doc: Any) -> None:
        """Закрытие документа без сохранения (двойная попытка)."""
        try:
            doc.Close(False)
        except Exception:
            try:
                doc.Close()
            except Exception as exc:
                self.log("warning", f"Не удалось закрыть документ: {exc}")
        time.sleep(0.5)

    def _get_doc3d(self, doc: Any) -> Any:
        """IKompasDocument3D (для TopPart); fallback — сам документ."""
        if self.conn.api7_module is not None:
            try:
                return self.conn.api7_module.IKompasDocument3D(doc)
            except Exception as exc:
                self.log("warning", f"IKompasDocument3D: {exc}; динамический fallback")
        return doc

    # ------------------------------------------------------------------

    def _walk_parts_inmemory(self, part: Any, parts: List[PartInfo],
                             refs: set, seen_paths: set,
                             depth: int, base_dir: str) -> None:
        """
        Рекурсивный обход компонентов сборки В ПАМЯТИ (проверено на v24,
        по образцу «Сводника»): part.Parts -> CastTo(IPart7) -> рекурсия
        по p7.Parts; подсборки с диска не открываются. Дедупликация:
          refs       — по Reference компонента (защита от циклов дерева);
          seen_paths — по пути файла детали (один файл = один экспорт).
        Путь файла: у IPart7 нет FilePath; свойства — PathName (обычно
        полный путь), FileName (имя файла), Path (каталог). Относительное
        имя достраиваем от Path либо от каталога сборки.
        TODO_KOMPAS: семантика PathName/FileName/Path подобрана эмпирически;
        при странностях смотреть лог строк «Компонент: ...».
        """
        if self._cancelled():
            return
        if depth > MAX_ASSEMBLY_DEPTH:
            self.log("warning", "Превышена глубина вложенности компонентов")
            return
        try:
            collection = part.Parts
        except Exception as exc:
            self.log("warning", f"part.Parts: {exc}")
            return
        children = com_items(collection)
        if not children:
            return

        for child in children:
            if self._cancelled():
                return
            try:
                p7 = CastTo(child, "IPart7")
            except Exception:
                p7 = child

            name = safe_str(p7, "Name") or "Без имени"
            ref = safe_str(p7, "Reference")
            if not ref:
                ref = f"{safe_str(p7, 'Marking')}|{name}|{depth}"
            if ref in refs:
                continue
            refs.add(ref)

            path_name = safe_str(p7, "PathName")
            file_only = safe_str(p7, "FileName")
            dir_path = safe_str(p7, "Path")
            file_name = ""
            for cand in (path_name, file_only):
                if cand and os.path.splitext(cand)[1].lower() in (".a3d",
                                                                 ".m3d"):
                    file_name = cand
                    break
            if not file_name:
                file_name = path_name or file_only
            if not file_name:
                self.log("warning", f"Компонент без пути — пропущен: {name}")
                continue
            if not os.path.isabs(file_name):
                file_name = os.path.join(dir_path or base_dir, file_name)
            child_ext = os.path.splitext(file_name)[1].lower()
            self.log("info", f"Компонент: {name} -> {file_name}")

            if child_ext == ".a3d":
                # Подсборка — рекурсия в памяти, без открытия с диска.
                self._walk_parts_inmemory(p7, parts, refs, seen_paths,
                                          depth + 1, base_dir)
            elif child_ext == ".m3d":
                key = os.path.normcase(os.path.abspath(file_name))
                if key in seen_paths:
                    continue
                seen_paths.add(key)
                if is_standard_part_path(file_name):
                    self.log("info",
                             f"Пропущено стандартное изделие: "
                             f"{os.path.basename(file_name)}")
                    continue
                material = safe_str(p7, "Material")
                parts.append(PartInfo(
                    file_path=os.path.abspath(file_name),
                    name=name,
                    marking=safe_str(p7, "Marking"),
                    material=material,
                    is_sheet=self._sheet_heuristic(material),
                ))
            else:
                self.log("warning",
                         f"Компонент неизвестного типа пропущен: {file_name}")

    def _sheet_heuristic(self, material: str) -> Optional[bool]:
        """
        Предварительная оценка «листовая?» по материалу (слово «лист»).
        Истинный тип определяет ISheetMetalBody при экспорте; звёздочка
        в GUI означает эвристику.
        """
        low = (material or "").lower()
        if not low:
            return None
        return "лист" in low

    # ------------------------------------------------------------------

    def _walk_file(self, path: str, parts: List[PartInfo],
                   seen: set, depth: int) -> None:
        """Рекурсивный обход файла (.a3d — сборка, .m3d — деталь)."""
        if self._cancelled():
            return
        if depth > MAX_ASSEMBLY_DEPTH:
            self.log("warning", f"Превышена глубина вложенности: {path}")
            return
        key = os.path.normcase(os.path.abspath(path))
        if key in seen:
            return
        seen.add(key)

        self.log("info", f"Обход: {os.path.basename(path)}")
        doc = self._open_document(path)
        if doc is None:
            self.log("error", f"Не удалось открыть: {path}")
            return
        try:
            doc_type = safe_int(doc, "DocumentType", 0)
            self.log("info", f"Тип документа: {doc_type}")

            ext = os.path.splitext(path)[1].lower()
            if ext == ".m3d":
                # Пользователь указал отдельную деталь вместо сборки.
                parts.append(PartInfo(file_path=os.path.abspath(path)))
                return

            # Сборка: IAssemblyDocument -> TopPart -> компоненты.
            asm = None
            if self.conn.api7_module is not None:
                try:
                    asm = self.conn.api7_module.IAssemblyDocument(doc)
                except Exception as exc:
                    self.log("warning", f"IAssemblyDocument: {exc}")
            top_part = None
            if asm is not None:
                try:
                    top_part = asm.TopPart
                except Exception as exc:
                    self.log("warning", f"asm.TopPart: {exc}")
            if top_part is None:
                try:
                    top_part = self._get_doc3d(doc).TopPart
                except Exception as exc:
                    self.log("warning", f"doc3d.TopPart: {exc}")
            if top_part is None:
                self.log("error", "Не удалось получить TopPart сборки")
                return

            base_dir = os.path.dirname(os.path.abspath(path))
            self.log("info", "Обход компонентов в памяти (IPart7.Parts)")
            self._walk_parts_inmemory(top_part, parts, set(), seen, 1,
                                      base_dir)
        finally:
            self._close_document(doc)


# ======================================================================
# ЭКСПОРТ ОДНОЙ ДЕТАЛИ В DXF
# ======================================================================

class PartExporter:
    """
    Полный цикл обработки одной детали:

      листовая:   открыть -> ISheetMetalBody -> Straighten=True -> rebuild ->
                  SaveAs(временный .m3d) -> ПАУЗА (проверка развертки) ->
                  Straighten=False -> закрыть -> фрагмент с видом из
                  временного файла -> ksSaveToDXF. Fallback: проекция
                  «#Развертка» от исходного файла.

      нелистовая: открыть -> ПАУЗА (пользователь вращает модель, вид
                  «DXF_Export» подхватывается при наличии) -> SaveAs(врем.
                  .m3d) -> закрыть -> фрагмент с видом из временного файла
                  -> ksSaveToDXF. Fallback: проекция «#Спереди».

    Документ создаётся как ФРАГМЕНТ (type=2): без рамки, штампа и оформления —
    в DXF остаётся только геометрия. Масштаб вида 1:1.
    """

    def __init__(self, conn: KompasConnection,
                 log: Callable[[str, str], None], tmp_dir: str,
                 template_path: Optional[str] = None):
        self.conn = conn
        self.log = log
        self.tmp_dir = tmp_dir
        # Шаблон фрагмента (.frw): если задан и существует — фрагменты
        # создаются из него (без лишних надписей/оформления по умолчанию).
        self.template_path = template_path or None

    # ------------------------------------------------------------------
    # Вспомогательные операции с документами
    # ------------------------------------------------------------------

    def _open_or_reuse(self, path: str) -> Tuple[Any, bool]:
        """
        Открыть деталь; если она уже открыта пользователем — использовать
        существующий документ (не закрывать после).
        Возвращает (doc, opened_by_us).
        """
        app = self.conn.application
        target = os.path.normcase(os.path.abspath(path))
        # Поиск среди открытых документов по PathName.
        try:
            docs = app.Documents
            count = com_count(docs) or 0
            for i in range(1, count + 1):   # TODO_KOMPAS: база индексации Documents
                doc = com_item(docs, i)
                if doc is None:
                    continue
                if os.path.normcase(safe_str(doc, "PathName")) == target:
                    self.log("info", "Документ уже открыт — использую существующий")
                    return doc, False
        except Exception as exc:
            self.log("warning", f"Перебор открытых документов не удался: {exc}")

        # Открытие с повторами (КОМПАС может быть занят модальным окном).
        for attempt in range(1, OPEN_RETRY_COUNT + 1):
            try:
                doc = app.Documents.Open(path)
                if doc is not None:
                    time.sleep(1.0)
                    return doc, True
            except Exception as exc:
                self.log("warning", f"Открытие не удалось (попытка {attempt}): {exc}")
            self.conn.pump(OPEN_RETRY_WAIT_S)
        return None, False

    def _close_doc(self, doc: Any, opened_by_us: bool) -> None:
        """Закрыть документ, если открывали мы. Пользовательский не трогаем."""
        if doc is None:
            return
        if not opened_by_us:
            self.log("info", "Документ был открыт пользователем — не закрываю")
            return
        try:
            doc.Close(False)
        except Exception:
            try:
                doc.Close()
            except Exception as exc:
                self.log("warning", f"Не удалось закрыть документ: {exc}")
        time.sleep(0.5)

    def _try_activate_saved_view(self, doc: Any, doc3d: Any, top: Any) -> bool:
        """
        Попытка найти и активировать сохранённый пользователем вид
        с именем «DXF_Export» (чтобы чертёж взял именно эту ориентацию).
        TODO_KOMPAS: реальные имена свойств/методов активации вида в 3D
        не документированы — перебираем кандидатов, всё в try/except.
        """
        candidates = []
        for getter in (
            lambda: doc.DocumentFrames,
            lambda: doc3d.Views,
            lambda: getattr(top, "Views", None),
            lambda: getattr(top, "DocumentFrames", None),
        ):
            try:
                coll = getter()
            except Exception:
                coll = None
            if coll is not None:
                candidates.append(coll)

        for coll in candidates:
            items = com_items(coll)
            for item in items:
                name = safe_str(item, "Name").strip()
                if name.lower() != SAVED_VIEW_NAME.lower():
                    continue
                self.log("info", f"Найден сохранённый вид «{name}» — активирую")
                for method in ("Activate", "SetCurrent", "SetActive"):
                    try:
                        getattr(item, method)()
                        self.log("info", f"Активация вида: {method}() — ок")
                        return True
                    except Exception:
                        continue
                self.log("warning", "Вид найден, но активировать не удалось")
        self.log("info", f"Сохранённый вид «{SAVED_VIEW_NAME}» не найден — "
                         "использую текущую ориентацию модели")
        return False

    # ------------------------------------------------------------------
    # 2D-конвейер: фрагмент -> ассоциативный вид -> DXF
    # ------------------------------------------------------------------

    def _activate_current_document(self) -> None:
        """Активировать текущий документ КОМПАС (окно фрагмента — на передний
        план, чтобы пользователь мог править вид на паузе)."""
        app = getattr(self.conn, "application", None)
        if app is None:
            return
        try:
            doc7 = app.ActiveDocument
        except Exception:
            doc7 = None
        if doc7 is None:
            return
        try:
            doc7.Activate()
        except Exception:
            pass

    def _close_api7_doc(self, doc7: Any) -> None:
        """Закрыть документ API7 без сохранения (уборка при отказе шаблона)."""
        if doc7 is None:
            return
        try:
            doc7.Close(False)
        except Exception:
            pass
        time.sleep(0.3)

    def _build_fragment_view(self, source_m3d: str, projection_name: str,
                             dxf_path: str,
                             pause: Optional[Callable[[str, str], None]] = None
                             ) -> Tuple[int, bool, str]:
        """
        Создаёт фрагмент 1:1, вставляет ассоциативный вид из 3D-модели
        и сохраняет в DXF. Проверенная последовательность (batch_export_dxf_v3).
        Если передан pause — после создания вида выполнение останавливается:
        пользователь в КОМПАС поправляет ориентацию вида и удаляет лишние
        надписи/обозначения, затем жмёт «Продолжить» — и только после этого
        фрагмент сохраняется в DXF в отредактированном виде.
        Возвращает (ссылка вида, успех, сообщение).
        """
        i_document2d = None
        try:
            kompas = self.conn.kompas_object

            # 1) Объект 2D-документа API5. Приоритет — шаблон пользователя
            #    (.frw): фрагмент создаётся из шаблона (API7
            #    AddNewDocumentFromTemplateEx), 2D-интерфейс — через
            #    ActiveDocument2D. Не вышло — обычный пустой фрагмент.
            if self.template_path:
                if os.path.isfile(self.template_path):
                    try:
                        doc7 = self.conn.application.Documents.\
                            AddNewDocumentFromTemplateEx(self.template_path,
                                                         True)
                        time.sleep(AFTER_CREATE_DOC_S)
                        i_document2d = kompas.ActiveDocument2D()
                        if i_document2d is not None:
                            self.log("info", "Фрагмент создан из шаблона: "
                                     + os.path.basename(self.template_path))
                        else:
                            self.log("warning",
                                     "Шаблон не дал 2D-интерфейс — создаю "
                                     "обычный фрагмент")
                            self._close_api7_doc(doc7)
                    except Exception as exc:
                        i_document2d = None
                        self.log("warning",
                                 f"Шаблон фрагмента не применился: {exc}")
                else:
                    self.log("warning", "Файл шаблона не найден: "
                             f"{self.template_path} — обычный фрагмент")

            if i_document2d is None:
                i_document2d = kompas.Document2D()
                if i_document2d is None:
                    return 0, False, "Document2D() вернул None"

                # 2) Параметры документа: ФРАГМЕНТ (без рамки и штампа).
                doc_param = kompas.GetParamStruct(self.conn.KO_DOCUMENT_PARAM)
                doc_param.Init()
                doc_param.type = DOC_TYPE_FRAGMENT
                if not i_document2d.ksCreateDocument(doc_param):
                    return 0, False, \
                        "ksCreateDocument вернул 0 (фрагмент не создан)"
                time.sleep(AFTER_CREATE_DOC_S)

            # 3) Параметры ассоциативного вида с модели.
            raw_avp = kompas.GetParamStruct(self.conn.KO_ASSOCIATION_VIEW_PARAM)
            if self.conn.api5_module is not None:
                avp = self.conn.api5_module.ksAssociationViewParam(raw_avp)
            else:
                avp = raw_avp  # динамический fallback
            avp.Init()
            avp.disassembly = False          # без разборки сборки
            avp.fileName = source_m3d        # источник геометрии (3D-модель)
            avp.hiddenLinesShow = False      # невидимые линии не показывать
            avp.hiddenLinesStyle = 4
            avp.projBodies = True            # проектировать тела
            avp.projectionLink = False       # БЕЗ ассоциативной связи с моделью
            # Имя проекции: '' — текущая ориентация модели (для временных
            # файлов), '#Развертка' / '#Спереди' — стандартные проекции.
            avp.projectionName = projection_name
            avp.projSurfaces = True
            avp.projThreads = True
            avp.sameHatch = False
            avp.section = False              # не разрез
            avp.tangentEdgesShow = False
            avp.tangentEdgesStyle = 2
            avp.visibleLinesStyle = 1

            # 4) Параметры самого вида: масштаб 1:1, без подписи.
            raw_vp = avp.GetViewParam()
            if self.conn.api5_module is not None:
                vp = self.conn.api5_module.ksViewParam(raw_vp)
            else:
                vp = raw_vp
            vp.Init()
            vp.angle = 0
            vp.color = 0
            vp.name = ""       # пустое имя — без подписи вида
            vp.scale_ = 1      # МАСШТАБ 1:1
            vp.state = 0
            vp.x = 0           # точка вставки — начало координат фрагмента
            vp.y = 0

            # 5) Создание вида на листе. Ненулевая ссылка = успех (проверено).
            view_ref = i_document2d.ksCreateSheetArbitraryView(avp, 0)
            time.sleep(AFTER_VIEW_S)
            if not view_ref:
                return 0, False, (
                    "Вид не создан (ksCreateSheetArbitraryView = 0); "
                    f"источник: {os.path.basename(source_m3d)}, "
                    f"проекция: «{projection_name or '(текущая ориентация)'}»")

            # 6) ПАУЗА: пользователь правит вид в КОМПАС (поворот, удаление
            #    лишних надписей/обозначений), затем продолжает — DXF
            #    сохраняется в отредактированном виде.
            if pause is not None:
                self._activate_current_document()
                self.log("info", "Пауза: проверка вида во фрагменте")
                pause("Проверьте вид во фрагменте",
                      f"Чертёж-фрагмент для «{os.path.basename(source_m3d)}» "
                      "создан в КОМПАС.\n"
                      "Поверните вид правильной стороной (выделите вид и "
                      "поверните его),\n"
                      "при необходимости удалите лишние надписи и обозначения "
                      "(сгибы, оси и т.п.).\n"
                      "Затем нажмите «Продолжить» — фрагмент будет сохранён "
                      "в DXF в текущем виде.\n"
                      "«Пропустить деталь» — без сохранения DXF.")

            # 7) Перестройка чертежа: после правок пользователя (поворот
            #    вида, удаление надписей) объекты чертежа нужно
            #    переформировировать — иначе DXF сохранится «как было».
            try:
                if not i_document2d.ksRebuildDocument():
                    self.log("warning", "ksRebuildDocument вернул False")
            except Exception as exc:
                self.log("warning", f"ksRebuildDocument: {exc}")
            time.sleep(0.3)

            # 8) Экспорт в DXF.
            # TODO_KOMPAS: ksSaveToDXF не принимает параметров; версия/единицы
            # DXF определяются настройками КОМПАС (Параметры -> Совместимость ->
            # форматы обмена). Требуемый профиль: ASCII, мм — настройте в КОМПАС.
            save_result = i_document2d.ksSaveToDXF(dxf_path)
            if save_result and os.path.isfile(dxf_path) \
                    and os.path.getsize(dxf_path) > 0:
                size = os.path.getsize(dxf_path)
                return view_ref, True, f"DXF создан ({size:,} байт)".replace(",", " ")
            return view_ref, False, \
                f"ksSaveToDXF вернул {save_result}, файл DXF не создан"

        except (CancelledByUser, SkippedByUser):
            raise  # «Отмена»/«Пропустить» на паузе — не ошибка, пробрасываем
        except Exception as exc:
            return 0, False, f"Ошибка при создании вида/DXF: {exc}"
        finally:
            # Фрагмент закрываем всегда, без сохранения.
            if i_document2d is not None:
                try:
                    i_document2d.ksCloseDocument()
                except Exception:
                    pass
                time.sleep(0.3)

    # ------------------------------------------------------------------
    # Главный метод: экспорт одной детали
    # ------------------------------------------------------------------

    def export_part(self, part: PartInfo, dxf_path: str,
                    pause: Callable[[str, str], None]) -> Tuple[str, str]:
        """
        Обработка одной детали. pause(title, instruction) — блокирующий
        вызов, бросает CancelledByUser/SkippedByUser.
        Возвращает ("OK" | "ERROR", сообщение).
        """
        self.log("info", "=" * 70)
        self.log("info", f"Деталь: {os.path.basename(part.file_path)}")

        # Убираем возможный старый DXF от прошлого запуска.
        if os.path.isfile(dxf_path):
            try:
                os.remove(dxf_path)
            except OSError as exc:
                return "ERROR", f"Не удалось удалить старый DXF ({exc})"

        state = {"ok": False}                 # для finally-уборки
        doc = None
        opened_by_us = False
        sheet_body = None
        top = None
        straighten_set = False
        temp_m3d: Optional[str] = None

        try:
            # --- A. Открытие детали и определение «листовая?» ---
            doc, opened_by_us = self._open_or_reuse(part.file_path)
            if doc is None:
                return "ERROR", "Не удалось открыть файл детали в КОМПАС"

            doc3d = self.conn.api7_module.IKompasDocument3D(doc) \
                if self.conn.api7_module is not None else doc
            top = doc3d.TopPart
            if top is None:
                return "ERROR", "Не удалось получить TopPart детали"

            # Cast к ISheetMetalBody: успех = листовая деталь (проверено
            # test_unfold_v2.py). Без api7-обёрток считаем деталь нелистовой.
            if self.conn.api7_module is not None:
                try:
                    sheet_body = self.conn.api7_module.ISheetMetalBody(top)
                except Exception:
                    sheet_body = None
            part.is_sheet = sheet_body is not None
            self.log("info", "Тип детали: "
                     + ("листовая (ISheetMetalBody)" if part.is_sheet
                        else "нелистовая"))

            stem = os.path.splitext(os.path.basename(part.file_path))[0]

            if sheet_body is not None:
                # --- B1. ЛИСТОВАЯ: развертка через выпрямление тела ---
                temp_m3d = os.path.join(
                    self.tmp_dir, sanitize_filename(stem) + "_unfold.m3d")

                # Выпрямляем тело листовой детали и перестраиваем модель.
                sheet_body.Straighten = True
                straighten_set = True
                top.RebuildModel(True)
                doc3d.RebuildDocument()
                time.sleep(REBUILD_WAIT_S)

                # Сохраняем развёрнутое состояние во временный файл ДО паузы —
                # фиксируем детерминированное состояние для вида.
                if os.path.isfile(temp_m3d):
                    os.remove(temp_m3d)
                doc.SaveAs(temp_m3d)
                if not (os.path.isfile(temp_m3d)
                        and os.path.getsize(temp_m3d) > 0):
                    return "ERROR", \
                        "Не удалось сохранить развёрнутую модель во временный файл"

                # ПАУЗА: пользователь проверяет развертку.
                pause("Развертка построена — проверьте",
                      f"Развертка для детали «{stem}» построена.\n"
                      "Проверьте её в окне КОМПАС; при необходимости "
                      "откорректируйте ориентацию модели.\n"
                      "Нажмите «Продолжить» для создания фрагмента и экспорта "
                      "в DXF,\nлибо «Пропустить деталь».")

                # Возвращаем детали исходную (гнутую) форму и закрываем.
                try:
                    sheet_body.Straighten = False
                    straighten_set = False
                    top.RebuildModel(True)
                except Exception as exc:
                    self.log("warning", f"Straighten=False: {exc}")
                self._close_doc(doc, opened_by_us)
                doc = None

                # Фрагмент из временной (развёрнутой) модели, текущая
                # ориентация (пустая projectionName) — проверено test_unfold_v2.
                _, ok, message = self._build_fragment_view(temp_m3d, "",
                                                           dxf_path, pause)
                if not ok:
                    # Fallback: проекция «#Развертка» от исходного файла
                    # (работает, если в модели построена плоская развертка).
                    self.log("warning",
                             "Вид из временной развертки не создан — пробую "
                             f"проекцию «{FLAT_PROJECTION_NAME}» от исходного файла")
                    _, ok, message = self._build_fragment_view(
                        part.file_path, FLAT_PROJECTION_NAME, dxf_path, pause)
                state["ok"] = ok
                return ("OK" if ok else "ERROR"), message

            # --- B2. НЕЛИСТОВАЯ: ориентацию задаёт пользователь ---
            # ПАУЗА: пользователь вращает модель / сохраняет вид «DXF_Export».
            pause("Поверните модель",
                  f"Поверните модель детали «{stem}» в КОМПАС нужной стороной.\n"
                  f"Рекомендуется сохранить текущий вид с именем "
                  f"«{SAVED_VIEW_NAME}» (команда «Сохранить вид»),\n"
                  "чтобы программа использовала именно эту ориентацию.\n"
                  "Затем нажмите «Продолжить».")

            # Подхватываем сохранённый вид, если пользователь его создал.
            self._try_activate_saved_view(doc, doc3d, top)

            # Фиксируем текущую ориентацию во временном файле; исходник не трогаем.
            # TODO_KOMPAS: механизм (SaveAs во врем. файл + пустая projectionName)
            # проверен для развертки; для произвольного поворота — fallback «#Спереди».
            temp_m3d = os.path.join(
                self.tmp_dir, sanitize_filename(stem) + "_view.m3d")
            if os.path.isfile(temp_m3d):
                os.remove(temp_m3d)
            doc.SaveAs(temp_m3d)
            if not (os.path.isfile(temp_m3d) and os.path.getsize(temp_m3d) > 0):
                return "ERROR", \
                    "Не удалось сохранить модель с выбранной ориентацией"
            self._close_doc(doc, opened_by_us)
            doc = None

            _, ok, message = self._build_fragment_view(temp_m3d, "", dxf_path)
            if not ok:
                self.log("warning",
                         "Вид из временной модели не создан — fallback на "
                         f"проекцию «{FRONT_PROJECTION_NAME}»")
                _, ok, message = self._build_fragment_view(
                    part.file_path, FRONT_PROJECTION_NAME, dxf_path, pause)
            state["ok"] = ok
            return ("OK" if ok else "ERROR"), message

        finally:
            # Уборка при любом исходе (в т.ч. отмена/пропуск на паузе):
            # вернуть детали гнутую форму, закрыть документы, удалить temp.
            if doc is not None:
                if straighten_set and sheet_body is not None:
                    try:
                        sheet_body.Straighten = False
                        top.RebuildModel(True)
                    except Exception:
                        pass
                self._close_doc(doc, opened_by_us)
            if temp_m3d and os.path.isfile(temp_m3d):
                if state["ok"]:
                    try:
                        os.remove(temp_m3d)
                    except OSError:
                        pass
                else:
                    self.log("warning",
                             f"Временный файл сохранён для диагностики: {temp_m3d}")


# ======================================================================
# РАБОЧИЙ ПОТОК (весь COM — только здесь)
# ======================================================================

class ExportWorker(threading.Thread):
    """
    Фоновый поток для работы с КОМПАС (обход сборки или экспорт деталей).
    GUI общается с потоком через:
      * msg_queue  — сообщения поток -> GUI (LOG/STATUS/PROGRESS/PAUSE/…);
      * события    — GUI -> поток (resume_event «Продолжить»,
                     skip_event «Пропустить деталь», cancel_event «Остановить»).
    Примечание: активный COM-вызов прервать нельзя — «Остановить» срабатывает
    после завершения текущего шага (открытие/построение/сохранение).
    """

    def __init__(self, msg_queue: "queue.Queue",
                 resume_event: threading.Event,
                 cancel_event: threading.Event,
                 skip_event: threading.Event,
                 mode: str,                       # "traverse" | "export"
                 payload: Any,                    # путь сборки | [(PartInfo, dxf)]
                 lic: LicenseManager,
                 logger: logging.Logger):
        super().__init__(daemon=True)
        self.msg_queue = msg_queue
        self.resume_event = resume_event
        self.cancel_event = cancel_event
        self.skip_event = skip_event
        self.mode = mode
        self.payload = payload
        self.lic = lic
        self.logger = logger
        self.template_path: Optional[str] = None   # шаблон фрагмента (.frw)

    # ---------------------- сообщения GUI ----------------------

    def _log(self, level: str, text: str) -> None:
        """Запись в файл-журнал и дублирование в GUI."""
        getattr(self.logger, level if level in ("info", "warning", "error")
                else "info")(text)
        self.msg_queue.put({"type": "LOG", "level": level, "text": text})

    def _status(self, text: str) -> None:
        self.msg_queue.put({"type": "STATUS", "text": text})

    # ---------------------- пауза ----------------------

    def _pause(self, title: str, instruction: str) -> None:
        """Блокирующая пауза до «Продолжить»; бросает Cancelled/Skipped."""
        self.skip_event.clear()
        self.resume_event.clear()
        self._status(title)
        self.msg_queue.put({"type": "PAUSE",
                            "title": title, "instruction": instruction})
        while not self.resume_event.wait(0.2):
            if self.cancel_event.is_set():
                raise CancelledByUser()
        if self.skip_event.is_set():
            raise SkippedByUser()

    # ---------------------- главный цикл ----------------------

    def run(self) -> None:
        tmp_dir: Optional[str] = None
        co_initialized = False
        try:
            if not PYWIN32_OK:
                raise KompasConnectionError(
                    "Не установлен пакет pywin32. Выполните: pip install pywin32")

            # COM инициализируется ВНУТРИ рабочего потока.
            pythoncom.CoInitialize()
            co_initialized = True

            self._log("info", "Подключаюсь к КОМПАС-3D…")
            conn = KompasConnection(self._log)
            conn.connect()
            self._log("info", "КОМПАС подключён ("
                      + ("запущен программой" if conn.started_by_us
                         else "подключено к работающему экземпляру") + ")")

            if self.mode == "traverse":
                # ---------- анализ сборки ----------
                walker = AssemblyWalker(conn, self._log, self.cancel_event)
                parts = walker.walk(str(self.payload))
                self.msg_queue.put(
                    {"type": "TRAVERSAL_DONE",
                     "parts": [p.to_dict() for p in parts]})
                return

            # ---------- экспорт выбранных деталей ----------
            jobs: List[Tuple[PartInfo, str]] = list(self.payload)
            tmp_dir = tempfile.mkdtemp(prefix="kompas_dxf_")
            self._log("info", f"Временный каталог: {tmp_dir}")
            exporter = PartExporter(conn, self._log, tmp_dir,
                                    self.template_path)

            total = len(jobs)
            ok = err = skip = 0
            cancelled = False
            stopped_by_limit = False
            had_errors = False
            self.msg_queue.put({"type": "PROGRESS", "done": 0, "total": total})

            for index, (part, dxf_path) in enumerate(jobs, start=1):
                if self.cancel_event.is_set():
                    cancelled = True
                    break
                self._status(f"[{index}/{total}] "
                             f"{part.name or os.path.basename(part.file_path)}")

                # Лицензия/триал: резерв слота до начала работы с деталью.
                if not self.lic.reserve():
                    self._log("warning",
                              "Лимит пробной версии исчерпан — экспорт "
                              "остановлен. " + LICENSE_CONTACT_HINT)
                    stopped_by_limit = True
                    break

                status, message = "ERROR", ""
                try:
                    status, message = exporter.export_part(part, dxf_path, self._pause)
                    if status == "OK":
                        self.lic.commit()      # успешный экспорт: счётчик +1
                    else:
                        self.lic.release()
                        had_errors = True
                except SkippedByUser:
                    self.lic.release()
                    status, message = "SKIP", "Пропущено пользователем"
                except CancelledByUser:
                    self.lic.release()
                    cancelled = True
                    self._status("Остановлено пользователем")
                    break
                except Exception as exc:  # непредвиденное — лог и дальше
                    self.lic.release()
                    self.logger.error(traceback.format_exc())
                    self._log("error", f"Непредвиденная ошибка: {exc}")
                    status, message = "ERROR", f"Непредвиденная ошибка: {exc}"
                    had_errors = True

                if status == "OK":
                    ok += 1
                    self._log("info", f"Результат: OK — {message}")
                elif status == "SKIP":
                    skip += 1
                    self._log("warning", f"Результат: SKIP — {message}")
                else:
                    err += 1
                    self._log("error", f"Результат: ERROR — {message}")

                self.msg_queue.put({"type": "PART_UPDATE",
                                    "file_path": part.file_path,
                                    "status": status,
                                    "message": message,
                                    "is_sheet": part.is_sheet})
                self.msg_queue.put({"type": "PROGRESS",
                                    "done": index, "total": total})

            lines = [f"Экспорт завершён.",
                     f"Успешно: {ok}",
                     f"С ошибками: {err}",
                     f"Пропущено: {skip}"]
            if cancelled:
                lines.append("Остановлено пользователем.")
            if stopped_by_limit:
                lines.append("Остановлено: лимит пробной версии исчерпан.\n"
                             + LICENSE_CONTACT_HINT)
            self._log("info", "\n".join(lines))
            self.msg_queue.put({"type": "DONE", "ok": ok, "err": err,
                                "skip": skip, "cancelled": cancelled,
                                "stopped_by_limit": stopped_by_limit,
                                "text": "\n".join(lines)})

            # Временные файлы: чистим при отсутствии ошибок, иначе оставляем
            # для диагностики (пути записаны в журнал).
            if tmp_dir and os.path.isdir(tmp_dir):
                if not had_errors:
                    shutil.rmtree(tmp_dir, ignore_errors=True)
                    tmp_dir = None

        except KompasConnectionError as exc:
            self.logger.error(traceback.format_exc())
            self._log("error", str(exc))
            self.msg_queue.put({"type": "FATAL", "text": str(exc)})
        except Exception as exc:
            self.logger.error(traceback.format_exc())
            self.msg_queue.put({"type": "FATAL",
                                "text": f"Критическая ошибка: {exc}"})
        finally:
            if tmp_dir and os.path.isdir(tmp_dir):
                self._log("info",
                          f"Временные файлы (диагностика) сохранены: {tmp_dir}")
            if co_initialized:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


# ======================================================================
# ГРАФИЧЕСКИЙ ИНТЕРФЕЙС (tkinter)
# ======================================================================

class App(tk.Tk):
    """Главное окно приложения."""

    def __init__(self, lic: LicenseManager, logger: logging.Logger):
        super().__init__()
        self.lic = lic
        self.logger = logger

        self.title(f"{APP_TITLE} — {VENDOR}")
        self.geometry("1020x780")
        self.minsize(880, 660)

        # Мосты GUI <-> рабочий поток.
        self.msg_queue: "queue.Queue" = queue.Queue()
        self.resume_event = threading.Event()
        self.cancel_event = threading.Event()
        self.skip_event = threading.Event()
        self.worker: Optional[ExportWorker] = None

        # Данные списка деталей.
        self.part_vars: List[Tuple[tk.BooleanVar, PartInfo]] = []
        self.rows: Dict[str, Dict[str, ttk.Label]] = {}   # normcase путь -> виджеты

        self._build_ui()
        self._set_buttons("idle")
        self._update_license_label()
        self._log_gui("info", STARTUP_INSTRUCTIONS)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self._poll_queue)

    # ------------------------------------------------------------------
    # Построение интерфейса
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        # ---- Верх: пути ----
        frm_top = ttk.Frame(self, padding=(8, 8, 8, 4))
        frm_top.pack(fill="x")

        self.asm_var = tk.StringVar()
        settings = load_settings()
        self.settings = settings
        self.out_var = tk.StringVar(
            value=settings.get("out_dir")
            or os.path.join(app_dir(), DEFAULT_OUT_SUBDIR))
        self.tpl_var = tk.StringVar(value=settings.get("template", ""))

        ttk.Label(frm_top, text="Сборка (.a3d):").grid(
            row=0, column=0, sticky="w", padx=(0, 4))
        ttk.Entry(frm_top, textvariable=self.asm_var).grid(
            row=0, column=1, sticky="we", padx=(0, 4))
        self.btn_browse_asm = ttk.Button(
            frm_top, text="Обзор…", command=self.on_browse_assembly)
        self.btn_browse_asm.grid(row=0, column=2, padx=(0, 4))
        self.btn_load = ttk.Button(
            frm_top, text="Загрузить детали", command=self.on_load_parts)
        self.btn_load.grid(row=0, column=3, padx=(0, 4))
        self.btn_add = ttk.Button(
            frm_top, text="Добавить файлы вручную", command=self.on_add_manual)
        self.btn_add.grid(row=0, column=4)

        ttk.Label(frm_top, text="Папка для DXF:").grid(
            row=1, column=0, sticky="w", padx=(0, 4), pady=(6, 0))
        ttk.Entry(frm_top, textvariable=self.out_var).grid(
            row=1, column=1, sticky="we", padx=(0, 4), pady=(6, 0))
        self.btn_browse_out = ttk.Button(
            frm_top, text="Обзор…", command=self.on_browse_out)
        self.btn_browse_out.grid(row=1, column=2, pady=(6, 0))

        ttk.Label(frm_top, text="Шаблон фрагмента:").grid(
            row=2, column=0, sticky="w", padx=(0, 4), pady=(6, 0))
        ttk.Entry(frm_top, textvariable=self.tpl_var).grid(
            row=2, column=1, sticky="we", padx=(0, 4), pady=(6, 0))
        self.btn_browse_tpl = ttk.Button(
            frm_top, text="Обзор…", command=self.on_browse_template)
        self.btn_browse_tpl.grid(row=2, column=2, pady=(6, 0))
        self.btn_clear_tpl = ttk.Button(
            frm_top, text="Сброс", command=self.on_clear_template)
        self.btn_clear_tpl.grid(row=2, column=3, pady=(6, 0))
        ttk.Label(frm_top, foreground="#606060",
                  text="необязательно: пустой фрагмент .frw —\n"
                       "надписи не придётся удалять вручную").grid(
                           row=2, column=4, sticky="w", pady=(6, 0))

        frm_top.columnconfigure(1, weight=1)

        # ---- Середина: список деталей ----
        frm_parts = ttk.LabelFrame(self, text="Детали", padding=(8, 4))
        frm_parts.pack(fill="both", expand=True, padx=8, pady=4)

        toolbar = ttk.Frame(frm_parts)
        toolbar.pack(fill="x", pady=(0, 4))
        ttk.Button(toolbar, text="Выбрать все",
                   command=lambda: self._toggle_all(True)).pack(side="left")
        ttk.Button(toolbar, text="Снять все",
                   command=lambda: self._toggle_all(False)).pack(
                       side="left", padx=(6, 0))
        self.counts_var = tk.StringVar(value="Найдено: 0    Выбрано: 0")
        ttk.Label(toolbar, textvariable=self.counts_var).pack(
            side="left", padx=16)
        ttk.Label(toolbar, text="(«листовая» без звёздочки — тип подтверждён "
                                "по API; до экспорта тип оценочный)").pack(
                                    side="right")

        container = ttk.Frame(frm_parts)
        container.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(container, highlightthickness=0)
        scroll = ttk.Scrollbar(container, orient="vertical",
                               command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self.inner_id = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw")
        self.inner.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")))
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(
                self.inner_id, width=e.width))
        # Колесо мыши — только при наведении на список.
        self.canvas.bind(
            "<Enter>",
            lambda e: self.canvas.bind_all(
                "<MouseWheel>",
                lambda ev: self.canvas.yview_scroll(
                    -1 * (ev.delta // 120), "units")))
        self.canvas.bind(
            "<Leave>",
            lambda e: self.canvas.unbind_all("<MouseWheel>"))

        # Заголовок колонок списка.
        header = ttk.Frame(self.inner)
        header.pack(fill="x", padx=2, pady=(0, 2))
        columns = [("", 4), ("Файл", 40), ("Тип", 13),
                   ("Обозначение", 18), ("Наименование", 24), ("Статус", 24)]
        for text, width in columns:
            ttk.Label(header, text=text, width=width, anchor="w",
                      font=("Segoe UI", 9, "bold")).pack(side="left")

        # ---- Статус и прогресс ----
        frm_status = ttk.LabelFrame(self, text="Статус", padding=(8, 4))
        frm_status.pack(fill="x", padx=8, pady=4)
        self.status_var = tk.StringVar(value="Готов к работе")
        ttk.Label(frm_status, textvariable=self.status_var,
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.instr_var = tk.StringVar(value="")
        ttk.Label(frm_status, textvariable=self.instr_var, justify="left",
                  wraplength=940, foreground="#0a3d91").pack(anchor="w")
        self.progress = ttk.Progressbar(frm_status, mode="determinate")
        self.progress.pack(fill="x", pady=(4, 0))

        # ---- Журнал ----
        frm_log = ttk.LabelFrame(self, text="Журнал", padding=(8, 4))
        frm_log.pack(fill="both", padx=8, pady=(4, 4))
        self.log_text = tk.Text(frm_log, height=9, state="disabled",
                                wrap="none", font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(frm_log, orient="vertical",
                                   command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)
        for tag, color in (("info", "#202020"), ("warning", "#b36b00"),
                           ("error", "#b00020"), ("ok", "#0a7a30")):
            self.log_text.tag_configure(tag, foreground=color)

        # ---- Кнопки управления ----
        frm_btn = ttk.Frame(self, padding=(8, 0, 8, 4))
        frm_btn.pack(fill="x")
        self.btn_export = ttk.Button(frm_btn, text="Начать экспорт",
                                     command=self.on_start_export)
        self.btn_export.pack(side="left")
        self.btn_continue = ttk.Button(frm_btn, text="Продолжить",
                                       command=self.on_continue)
        self.btn_continue.pack(side="left", padx=(8, 0))
        self.btn_skip = ttk.Button(frm_btn, text="Пропустить деталь",
                                   command=self.on_skip)
        self.btn_skip.pack(side="left", padx=(8, 0))
        self.btn_cancel = ttk.Button(frm_btn, text="Остановить",
                                     command=self.on_cancel)
        self.btn_cancel.pack(side="right")

        # ---- Нижняя строка: лицензия ----
        frm_bottom = ttk.Frame(self, padding=(8, 0, 8, 6))
        frm_bottom.pack(fill="x")
        ttk.Button(frm_bottom, text="Лицензия…",
                   command=self.show_license_dialog).pack(side="left")
        self.license_var = tk.StringVar(value="")
        self.license_label = ttk.Label(frm_bottom, textvariable=self.license_var)
        self.license_label.pack(side="left", padx=12)
        ttk.Label(frm_bottom,
                  text=f"{VENDOR} — {SUPPORT_EMAIL}").pack(side="right")

    # ------------------------------------------------------------------
    # Управление состоянием кнопок
    # ------------------------------------------------------------------

    def _set_buttons(self, state: str) -> None:
        """idle — простой; working — идёт работа; paused — ждём пользователя."""
        idle = state == "idle"
        working = state == "working"
        paused = state == "paused"
        for btn in (self.btn_load, self.btn_add, self.btn_browse_asm,
                    self.btn_browse_out, self.btn_export):
            btn.configure(state="normal" if idle else "disabled")
        self.btn_cancel.configure(
            state="normal" if (working or paused) else "disabled")
        self.btn_continue.configure(state="normal" if paused else "disabled")
        self.btn_skip.configure(state="normal" if paused else "disabled")

    def _worker_alive(self) -> bool:
        return bool(self.worker and self.worker.is_alive())

    # ------------------------------------------------------------------
    # Журнал и статус
    # ------------------------------------------------------------------

    def _log_gui(self, level: str, text: str) -> None:
        tag = level if level in ("info", "warning", "error", "ok") else "info"
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _update_license_label(self) -> None:
        self.license_var.set(self.lic.status_text())
        if self.lic.licensed:
            color = "#0a7a30"
        elif self.lic.corrupted:
            color = "#b00020"
        else:
            color = "#b36b00"
        self.license_label.configure(foreground=color)

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
        """Сохранить текущие настройки GUI (шаблон, папка вывода)."""
        self.settings["template"] = self.tpl_var.get().strip()
        self.settings["out_dir"] = self.out_var.get().strip()
        save_settings(self.settings)

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
        for widget in self.inner.winfo_children():
            if widget is not getattr(self, "_list_header", None):
                widget.destroy()
        # Заголовок колонок создаётся заново (он был уничтожен вместе с детьми).
        self._make_list_header()
        self.rows.clear()
        self.part_vars.clear()
        ordered = sorted(parts, key=lambda p: (p.marking, p.name, p.file_path))
        for part in ordered:
            # Нелистовые детали (по эвристике материала) — сняты с выбора;
            # неопределённый тип остаётся выбранным.
            self._append_part_row(part, selected=part.is_sheet is not False)

    def _make_list_header(self) -> None:
        header = ttk.Frame(self.inner)
        header.pack(fill="x", padx=2, pady=(0, 2))
        columns = [("", 4), ("Файл", 40), ("Тип", 13),
                   ("Обозначение", 18), ("Наименование", 24), ("Статус", 24)]
        for text, width in columns:
            ttk.Label(header, text=text, width=width, anchor="w",
                      font=("Segoe UI", 9, "bold")).pack(side="left")

    def _append_part_row(self, part: PartInfo, selected: bool = True) -> None:
        """Добавление одной строки в список деталей."""
        var = tk.BooleanVar(value=selected)
        row = ttk.Frame(self.inner)
        row.pack(fill="x", padx=2, pady=1)
        ttk.Checkbutton(row, variable=var,
                        command=self._update_counts).grid(
                            row=0, column=0, sticky="w")
        ttk.Label(row, text=os.path.basename(part.file_path),
                  width=40, anchor="w").grid(row=0, column=1, sticky="w")
        type_lbl = ttk.Label(row, text=part_type_text(part.is_sheet),
                             width=13, anchor="w")
        type_lbl.grid(row=0, column=2, sticky="w")
        ttk.Label(row, text=part.marking, width=18,
                  anchor="w").grid(row=0, column=3, sticky="w")
        ttk.Label(row, text=part.name, width=24,
                  anchor="w").grid(row=0, column=4, sticky="w")
        status_lbl = ttk.Label(row, text="—", width=24, anchor="w")
        status_lbl.grid(row=0, column=5, sticky="w")
        self.rows[os.path.normcase(part.file_path)] = {
            "status": status_lbl, "type": type_lbl}
        self.part_vars.append((var, part))

    def _toggle_all(self, value: bool) -> None:
        for var, _ in self.part_vars:
            var.set(value)
        self._update_counts()

    def _update_counts(self) -> None:
        total = len(self.part_vars)
        chosen = sum(1 for var, _ in self.part_vars if var.get())
        self.counts_var.set(f"Найдено: {total}    Выбрано: {chosen}")

    def _update_part_row(self, file_path: str, status: str,
                         message: str, is_sheet: Optional[bool]) -> None:
        row = self.rows.get(os.path.normcase(file_path))
        if row is None:
            return
        text = {"OK": "OK", "ERROR": "ОШИБКА", "SKIP": "ПРОПУСК"}.get(status, "—")
        if message:
            text += f": {message}"
        row["status"].configure(text=text[:120])
        row["status"].configure(foreground={
            "OK": "#0a7a30", "ERROR": "#b00020",
            "SKIP": "#b36b00"}.get(status, "#202020"))
        if is_sheet is not None:
            row["type"].configure(text=part_type_text(is_sheet))

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

        # Расчёт путей DXF с разрешением коллизий имён в рамках запуска.
        used_names = set()
        jobs: List[Tuple[PartInfo, str]] = []
        for part in selected:
            base = build_dxf_basename(part)
            name = base
            suffix = 2
            while name.lower() in used_names:
                name = f"{base} ({suffix})"
                suffix += 1
            used_names.add(name.lower())
            jobs.append((part, os.path.join(out_dir, name + ".dxf")))

        self._log_gui("info", f"Начинаю экспорт: {len(jobs)} дет., папка: {out_dir}")
        self.progress.configure(value=0, maximum=len(jobs))
        self._start_worker("export", jobs)

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
            self.progress.configure(maximum=msg.get("total", 100),
                                    value=msg.get("done", 0))

        elif mtype == "PAUSE":
            # Пауза: показать инструкцию, включить «Продолжить»/«Пропустить».
            self._set_buttons("paused")
            self.instr_var.set(
                f"{msg.get('title', '')}\n\n{msg.get('instruction', '')}\n\n"
                ">>> ЖДЁМ ВАШЕГО ДЕЙСТВИЯ <<<")

        elif mtype == "PART_UPDATE":
            self._update_part_row(msg.get("file_path", ""),
                                  msg.get("status", "—"),
                                  msg.get("message", ""),
                                  msg.get("is_sheet"))
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
    # Диалог лицензии
    # ------------------------------------------------------------------

    def show_license_dialog(self) -> None:
        """Окно лицензии: статус, HWID для заказа, подключение license.key."""
        win = tk.Toplevel(self)
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
                  font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=12,
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


# ======================================================================
# ТОЧКА ВХОДА
# ======================================================================

def main() -> None:
    logger = setup_logging()
    logger.info("=" * 60)
    logger.info("Запуск kompas_dxf_exporter")

    lic = LicenseManager()
    try:
        lic.load()
        logger.info(f"HWID: {lic.get_hwid()}; лицензия: "
                    f"{'активна (' + lic.customer + ')' if lic.licensed else 'нет'}; "
                    f"использовано триала: {lic.used}"
                    + ("; СОСТОЯНИЕ ПОВРЕЖДЕНО" if lic.corrupted else ""))
    except Exception:
        logger.error(traceback.format_exc())

    app = App(lic, logger)
    app.mainloop()


if __name__ == "__main__":
    main()
