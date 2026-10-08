# -*- coding: utf-8 -*-
"""Собирает index.html (самостоятельная страница для GitHub Pages / открытия с диска)
из web/sverka.html — исходника, который публикуется по ссылке в Claude.
Запуск:  python web/build_index.py"""
from pathlib import Path

root = Path(__file__).resolve().parent.parent
body = (root / 'web' / 'sverka.html').read_text(encoding='utf-8')
head = ('<!doctype html><html lang="ru"><head><meta charset="utf-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">'
        '<link rel="icon" type="image/svg+xml" href="static/favicon.svg"></head><body>\n')
(root / 'index.html').write_text(head + body + '\n</body></html>\n', encoding='utf-8')
print('index.html обновлён')
