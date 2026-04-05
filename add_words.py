import argparse
import sys
from collections import OrderedDict
from pathlib import Path

from data_manager import DataManager
from parser import parse_input_file
from sheets_handler import GoogleSheetsHandler
from utils import load_config


def configure_console_output() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")



def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Добавление слов из input.txt в Google Sheets"
    )
    parser.add_argument(
        "--kapitel",
        type=int,
        required=True,
        metavar="N",
        help=(
            "Номер Kapitel из учебника для листа 📗 Fokus Deutsch и для 🗣 REDEMITTEL Sprechen. "
            "Лист ✍️ REDEMITTEL Schreiben использует только блоки «Schreiben B2 - Teil n»; этот номер на него не влияет."
        ),
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=Path("input.txt"),
        help="Путь к входному файлу со словами",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("config.json"),
        help="Путь к JSON-конфигу проекта",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Показать план действий без записи в Google Sheets",
    )
    return parser



def main() -> None:
    configure_console_output()
    args = build_argument_parser().parse_args()

    print("Загрузка конфигурации...")
    config = load_config(args.config)
    print("Чтение входного файла...")
    categories_words = parse_input_file(args.file)

    print("Найдены категории во входном файле:")
    for category, words in categories_words.items():
        print(f"- {category}: {len(words)} слов")

    gs_handler = GoogleSheetsHandler.from_config(config, dry_run=args.dry_run)
    data_manager = DataManager(gs_handler, config=config)

    print("Подключение к Google Sheets...")
    gs_handler.connect()
    print(f"Подключение успешно. Доступно листов: {len(gs_handler.sheet_title_to_id)}")

    words_by_sheet: OrderedDict[str, list[tuple[str, str]]] = OrderedDict()
    for category, words in categories_words.items():
        sheet_name = data_manager._resolve_sheet_name(category)
        words_by_sheet.setdefault(sheet_name, []).extend(words)

    results = []
    for sheet_name, words in words_by_sheet.items():
        print(f"Обработка листа: {sheet_name} ({len(words)} слов)")
        result = data_manager.add_words_to_sheet(
            sheet_name=sheet_name,
            words=words,
            kapitel=args.kapitel,
        )
        results.append(result)

    print(f"Обработка завершена. Dry run: {'да' if args.dry_run else 'нет'}")
    for result in results:
        print(
            f"[{result.sheet_name}] новых: {result.added_count}, "
            f"дубликатов: {result.duplicate_count}, "
            f"обновлений приоритета: {result.priority_updated_count}"
        )
        planned_actions = getattr(result, "planned_actions", None)
        if args.dry_run and planned_actions:
            print(f"  План для [{result.sheet_name}]:")
            for action in planned_actions:
                if action.action == "create_teil":
                    print(f"  - создать часть {action.teil} ({action.details})")
                elif action.action == "fill_slot":
                    print(f"  - записать '{action.term}' в строку {action.row_number} части {action.teil}")
                elif action.action == "fill_new_teil":
                    print(f"  - записать '{action.term}' в новую часть {action.teil}, строка {action.row_number}")
                elif action.action == "priority_update":
                    print(f"  - обновить приоритет '{action.term}' в строке {action.row_number} ({action.details})")


if __name__ == "__main__":
    main()
