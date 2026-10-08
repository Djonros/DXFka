# -*- coding: utf-8 -*-
"""Лицензирование: HWID, состояние пробной версии, файл лицензии."""

import hashlib
import hmac
import json
import os
import re
import shutil
import threading
import uuid
import winreg
from typing import Any, Dict, Optional, Tuple

from dxfka.config import (
    LICENSE_FILE, PRODUCT_ID, REG_KEY, SECRET, STATE_DIR, STATE_FILE,
    SUPPORT_EMAIL, TRIAL_EXPORT_LIMIT, VENDOR)
from dxfka.utils import app_dir

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
            if self.licensed and self.corrupted:
                # С действующей лицензией счётчик триала не нужен:
                # переподписываем его текущим секретом, чтобы флаг
                # «повреждено» не висел в журнале и не вернулся позже.
                self.corrupted = False
                self._write_state()
                self._write_install_marker()

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
