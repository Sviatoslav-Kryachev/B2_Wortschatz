import re
from pathlib import Path

from utils import repair_mojibake


KNOWN_CATEGORY_PREFIXES = (
    "Существительное",
    "Глаголы",
    "Наречия / Прилагательные",
    "Фраза:",
    "✍️ REDEMITTEL Schreiben",
)


def _is_category_line(line: str) -> bool:
    normalized = line.strip()
    if not normalized or _line_has_term_definition_separator(normalized):
        return False
    if normalized == "Фраза:":
        return False
    return any(normalized.startswith(prefix) for prefix in KNOWN_CATEGORY_PREFIXES)


def _split_term_definition(line: str) -> tuple[str, str] | None:
    """Термин и перевод: em dash, en dash или «пробел-дефис-пробел»."""
    for sep in ("\u2014", "\u2013"):  # — –
        if sep in line:
            term, _, definition = line.partition(sep)
            term, definition = term.strip(), definition.strip()
            if term and definition:
                return term, definition
    if " - " in line:
        term, _, definition = line.partition(" - ")
        term, definition = term.strip(), definition.strip()
        if term and definition:
            return term, definition
    return None


def _line_has_term_definition_separator(line: str) -> bool:
    if "\u2014" in line or "\u2013" in line:
        return True
    return bool(re.search(r"\S\s+-\s+\S", line))


def _normalize_category_name(line: str) -> str:
    normalized = line.strip()
    if normalized.endswith(":"):
        normalized = normalized[:-1].strip()
    return normalized


def parse_input_file(file_path: str | Path) -> dict[str, list[tuple[str, str]]]:
    """
    Возвращает словарь вида:
    {
        "Существительное": [("die Voraussetzung (-en)", "предпосылка")],
        "Фраза: 🗣 REDEMITTEL": [...],
        "✍️ REDEMITTEL Schreiben": [...],
        ...
    }
    """
    categories: dict[str, list[tuple[str, str]]] = {}
    current_category: str | None = None

    with Path(file_path).open("r", encoding="utf-8") as source:
        for raw_line in source:
            line = repair_mojibake(raw_line).strip()

            if not line or line.startswith("="):
                continue

            if _is_category_line(line):
                next_category = _normalize_category_name(line)
                if next_category:
                    current_category = next_category
                    categories.setdefault(current_category, [])
                continue

            pair = _split_term_definition(line)
            if pair is None:
                continue

            if current_category is None:
                raise ValueError(
                    "Найдена строка со словом до объявления категории: "
                    f"{line}"
                )

            term, definition = pair
            categories[current_category].append((term, definition))

    return categories
