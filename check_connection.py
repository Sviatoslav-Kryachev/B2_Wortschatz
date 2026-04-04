import argparse
import sys
from pathlib import Path

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
        description="Проверка подключения к Google Sheets"
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=Path("config.json"),
        help="Путь к JSON-конфигу проекта",
    )
    return parser


def main() -> None:
    configure_console_output()
    args = build_argument_parser().parse_args()

    config = load_config(args.config)
    gs_handler = GoogleSheetsHandler.from_config(config, dry_run=True)

    print("Проверка подключения к Google Sheets...")
    print(f"Config: {args.config}")
    print(f"Sheet ID: {gs_handler.sheet_id}")
    print(f"Credentials file: {gs_handler.credentials_file}")

    gs_handler.connect()

    print("Подключение успешно.")
    print(f"Найдено листов: {len(gs_handler.sheet_title_to_id)}")
    for title in gs_handler.sheet_title_to_id:
        print(f"- {title}")


if __name__ == "__main__":
    main()
