# -*- coding: utf-8 -*-
"""
Шаблон секрета лицензирования DXFka.

Сгенерируйте свои половины (A и B — случайные 32 байта, SECRET = A XOR B):

    python -c "import secrets; a=secrets.token_bytes(32); b=secrets.token_bytes(32); s=bytes(x^y for x,y in zip(a,b)); print('A:',a.hex()); print('B:',b.hex()); print('check:', s.hex())"

…объедините так, чтобы A XOR B давало ваш секрет s, сохраните как
license_secret.py рядом с kompas_dxf_exporter.py и keygen.py.
Файл license_secret.py добавлен в .gitignore и в репозиторий не входит.

Без этого файла программа работает в режиме разработки: пробная версия
считается нормально, но подключение лицензий недоступно (подпись
проверить нечем). Коммерческие сборки exe собираются с настоящим секретом.
"""

_SK_A = "00" * 32
_SK_B = "00" * 32

SECRET = bytes(a ^ b for a, b in zip(bytes.fromhex(_SK_A), bytes.fromhex(_SK_B)))
