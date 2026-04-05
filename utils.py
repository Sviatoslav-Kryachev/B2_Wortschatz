import json
import re
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


def term_duplicate_key(value: str) -> str:
    """
    Ключ для сравнения дубликатов: нормализует текст, удаляет:
    - формы глаголов (; ... ; hat ...)
    - множественное число в скобках (-en), (-n)
    - запятые с окончаниями , -en, , -n
    - возвратное sich
    - артикли в начале
    - пробелы и знаки препинания
    """
    base = normalize_text(value)
    
    # Удаляем возвратное "sich" в начале
    base = re.sub(r'^sich\s+', '', base)
    
    # Удаляем артикли в начале (der, die, das)
    base = re.sub(r'^(der|die|das)\s+', '', base)
    
    # Удаляем формы глаголов: ; что-то; hat что-то
    base = re.sub(r'\s*;.*$', '', base)
    
    # Удаляем скобки с окончаниями в конце (-en), (-n) и т.д.
    base = re.sub(r'\s*\([^)]*\)\s*$', '', base)
    
    # Удаляем запятые с окончаниями , -en, , -n
    base = re.sub(r'\s*,\s*-\s*[en]+\s*$', '', base)
    base = re.sub(r'\s*,\s*-[en]+\s*$', '', base)
    
    # Удаляем одиночные запятые в конце
    base = re.sub(r'\s*,\s*$', '', base)
    
    # Удаляем дефисы и лишние пробелы
    base = re.sub(r'\s*-\s*', ' ', base)
    base = re.sub(r'[,]', '', base)
    base = re.sub(r'\s+', ' ', base).strip()
    
    return base