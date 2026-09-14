#!/usr/bin/env python3
"""Демо-парсер «scraper»: сбор записей с сайта-каталога в Excel.

Примеры:
    python run.py --source demo --out demo_result.xlsx
    python run.py --source books --limit 5
    python run.py --source https://books.toscrape.com/ --out result.xlsx --json result.json
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import scraper_core as core


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="run.py",
        description="Собирает записи с сайта-каталога и сохраняет в Excel "
        "(листы «Данные» и «Сводка»).",
        epilog="Пресеты источников: demo (офлайн-каталог), books (books.toscrape.com).",
    )
    parser.add_argument(
        "--source",
        required=True,
        help="URL http(s)://, имя пресета (demo, books) или путь к HTML-файлу",
    )
    parser.add_argument(
        "--out", default="result.xlsx", help="куда сохранить Excel (по умолчанию result.xlsx)"
    )
    parser.add_argument("--limit", type=int, default=None, metavar="N", help="взять первые N записей")
    parser.add_argument(
        "--json", dest="json_out", metavar="FILE", help="дополнительно сохранить JSON"
    )
    parser.add_argument(
        "--timeout", type=int, default=None, metavar="SEC", help="таймаут сети в секундах"
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")

    args = build_arg_parser().parse_args(argv)

    if args.limit is not None and args.limit < 0:
        print("[ОШИБКА] --limit не может быть отрицательным.", file=sys.stderr)
        return 1

    try:
        html, base_url, label = core.load_source(args.source, timeout=args.timeout)
        records = core.parse_records(html, base_url)
    except core.ScrapingError as exc:
        print(f"[ОШИБКА] {exc}", file=sys.stderr)
        return exc.exit_code
    except Exception as exc:  # непредвиденное — короткое сообщение без портянки
        print(
            f"[ОШИБКА] Непредвиденная ошибка: {exc.__class__.__name__}: {exc}",
            file=sys.stderr,
        )
        return 1

    total_found = len(records)
    if args.limit is not None:
        records = records[: args.limit]

    try:
        core.build_workbook(records, Path(args.out), source_label=label)
        if args.json_out:
            core.write_json(records, Path(args.json_out), source_label=label)
    except OSError as exc:
        print(f"[ОШИБКА] Не удалось записать файл: {exc}", file=sys.stderr)
        return 1

    print(f"Источник: {label}")
    if args.limit is not None and total_found > len(records):
        print(f"Записей получено: {total_found}, взято первых {len(records)} (--limit {args.limit})")
    else:
        print(f"Записей получено: {len(records)}")
    print(f"Excel: {Path(args.out)} (листы «Данные» и «Сводка»)")
    if args.json_out:
        print(f"JSON:  {Path(args.json_out)}")
    if not records:
        print("[ВНИМАНИЕ] Источник не дал записей: в Excel записаны только заголовки и нулевая сводка.")
    print("Готово.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
