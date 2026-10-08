# -*- coding: utf-8 -*-
"""
DXFka — пакетный экспорт DXF из деталей сборки КОМПАС-3D v24.
Автор и поставщик: Djonros (djonros@gmail.com).

Точка входа. Код приложения — в пакете dxfka/:
  config.py            константы приложения, лицензирования, КОМПАС API
  compat.py            pywin32 (может отсутствовать)
  utils.py             утилиты, журнал, настройки, COM-коллекции
  licensing.py         HWID, пробная версия, файл лицензии
  models.py            PartInfo и исключения потока экспорта
  kompas/connection.py подключение к КОМПАС, сброс кеша COM
  kompas/walker.py     обход сборки
  kompas/exporter.py   экспорт детали в DXF
  worker.py            рабочий поток (весь COM — только здесь)
  ui/                  графический интерфейс (tkinter)

ЗАПУСК (Python 3.10+, Windows 10/11):
  pip install pywin32
  python kompas_dxf_exporter.py      (или kompas_dxf_exporter.pyw без консоли)
"""

from dxfka.main import main

if __name__ == "__main__":
    main()
