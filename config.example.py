"""Пример конфига. Скопируйте в config.py (он не публикуется и не коммитится).

Без config.py всё работает на значениях по умолчанию из scraper_core.py.
Секреты и токены держите только в config.py или переменных окружения —
никогда не хардкодьте их в коде и не коммитьте.
"""

import os

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)

TIMEOUT = 20  # секунды на запрос

# Пример чтения ключа из окружения (пригодится для источников с API):
# API_KEY = os.environ.get("SCRAPER_API_KEY", "")
