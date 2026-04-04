import json
from pathlib import Path


def repair_mojibake(value: str) -> str:
    suspicious_markers = ("Ð", "Ñ", "â", "ï", "ð")
    if not isinstance(value, str) or not any(marker in value for marker in suspicious_markers):
        return value

    for source_encoding in ("latin-1", "cp1252"):
        try:
            repaired = value.encode(source_encoding).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if repaired != value:
            return repaired

    return value


def repair_mojibake_in_data(value):
    if isinstance(value, str):
        return repair_mojibake(value)
    if isinstance(value, list):
        return [repair_mojibake_in_data(item) for item in value]
    if isinstance(value, dict):
        return {
            repair_mojibake_in_data(key): repair_mojibake_in_data(item)
            for key, item in value.items()
        }
    return value


def load_config(config_path: str | Path) -> dict:
    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл конфигурации не найден: {path}")

    raw_text = path.read_text(encoding="utf-8").strip()
    if not raw_text:
        raise ValueError(
            "config.json пустой. Заполни настройки Google Sheets перед запуском."
        )

    return repair_mojibake_in_data(json.loads(raw_text))


def normalize_text(value: str) -> str:
    return " ".join(value.strip().casefold().split())

