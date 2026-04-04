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
    if line.endswith(":"):
        return True
    return any(line.startswith(prefix) for prefix in KNOWN_CATEGORY_PREFIXES)


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
                current_category = _normalize_category_name(line)
                categories.setdefault(current_category, [])
                continue

            if "—" not in line:
                continue

            if current_category is None:
                raise ValueError(
                    "Найдена строка со словом до объявления категории: "
                    f"{line}"
                )

            term, definition = line.split("—", 1)
            categories[current_category].append((term.strip(), definition.strip()))

    return categories
