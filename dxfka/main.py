# -*- coding: utf-8 -*-
"""Точка входа приложения."""

import traceback

from dxfka.licensing import LicenseManager
from dxfka.ui.app import App
from dxfka.utils import setup_logging

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
