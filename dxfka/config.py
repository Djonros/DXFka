# -*- coding: utf-8 -*-
"""Константы приложения, лицензирования и КОМПАС API."""

import os
import sys
import tempfile

# ======================================================================
# КОНСТАНТЫ: приложение, брендинг
# ======================================================================

APP_VERSION = "2.0"
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
    DEV_MODE = False
except ImportError:
    SECRET = b"DXFka-dev-fallback-secret"
    DEV_MODE = True

# Счётчик триала, подписанный dev-секретом, не проходит проверку в сборке
# с настоящим секретом (и наоборот) — состояние выглядело «повреждённым».
# Поэтому режим разработки хранит своё состояние отдельно.
if DEV_MODE:
    STATE_FILE = "state.dev.json"
    REG_KEY += r"\Dev"

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
# Проекции для плоской нелистовой детали в авто-режиме: берётся та,
# где контур занимает наибольшую площадь (лицевая сторона пластины).
PLATE_PROJECTIONS = ("#Спереди", "#Сверху", "#Слева")
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

# Каталог приложения: для exe — папка с exe, иначе корень репозитория
# (на уровень выше пакета dxfka).
APP_DIR = (os.path.dirname(sys.executable) if getattr(sys, "frozen", False)
           else os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
LOG_FILE = os.path.join(APP_DIR, "kompas_dxf_exporter.log")
DEFAULT_OUT_SUBDIR = "DXF_Out"

STARTUP_INSTRUCTIONS = (
    "Порядок работы:\n"
    "  1. Укажите файл сборки (.a3d) и папку для DXF.\n"
    "  2. Нажмите «Загрузить детали» — программа проанализирует сборку:\n"
    "     будут найдены все детали, включая исполнения и локальные (встроенные\n"
    "     в сборку, без своего файла). Не хватает — «Добавить файлы вручную».\n"
    "  3. Выберите режим в меню «Обработка»: «Автоматически» или «Вручную».\n"
    "  4. Отметьте нужные детали (флажок в первой колонке) и нажмите\n"
    "     «Начать экспорт».\n"
    "  5. Вручную — на ПЕРВОЙ паузе каждой детали определите вид для резки\n"
    "     (листовая — проверьте развертку; нелистовая — поверните модель,\n"
    "     при желании сохраните вид «DXF_Export»), затем укажите ТОЛЩИНУ\n"
    "     (мм) — DXF сохранится в подпапку с этим номером (создаётся сама).\n"
    "     Автоматически — толщина из листового тела или из материала\n"
    "     («Лист 8 …»), паузы пропускаются; остановка — только на деталях\n"
    "     без толщины.\n"
    "  6. Результат: DXF 1:1 в мм, только контур (без рамки и надписей).\n"
    f"  Поддержка и лицензии: {VENDOR} — {SUPPORT_EMAIL}"
)
