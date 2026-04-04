from dataclasses import dataclass

from sheets_handler import GoogleSheetsHandler, PreparedWord


@dataclass(slots=True)
class CategoryProcessResult:
    sheet_name: str
    added_count: int
    duplicate_count: int
    priority_updated_count: int
    planned_actions: list | None = None


class DataManager:
    def __init__(self, sheets_handler: GoogleSheetsHandler, config: dict) -> None:
        self.sheets_handler = sheets_handler
        self.config = config
        self.sheet_layout = config.get("sheet_layout", {})

    def add_words_to_category(
        self,
        category: str,
        words: list[tuple[str, str]],
        kapitel: int,
    ) -> CategoryProcessResult:
        default_priority = int(self.sheet_layout.get("default_priority", 1))
        sheet_name = self._resolve_sheet_name(category)
        prepared_words = [
            PreparedWord(term=term, definition=definition, priority=default_priority)
            for term, definition in words
        ]
        sync_result = self.sheets_handler.sync_category_words(
            sheet_name=sheet_name,
            category=category,
            kapitel=kapitel,
            words=prepared_words,
        )
        return CategoryProcessResult(
            sheet_name=sheet_name,
            added_count=sync_result.added_count,
            duplicate_count=sync_result.duplicate_count,
            priority_updated_count=sync_result.priority_updated_count,
            planned_actions=sync_result.planned_actions,
        )

    def add_words_to_sheet(
        self,
        sheet_name: str,
        words: list[tuple[str, str]],
        kapitel: int,
    ) -> CategoryProcessResult:
        default_priority = int(self.sheet_layout.get("default_priority", 1))
        prepared_words = [
            PreparedWord(term=term, definition=definition, priority=default_priority)
            for term, definition in words
        ]
        sync_result = self.sheets_handler.sync_category_words(
            sheet_name=sheet_name,
            category=sheet_name,
            kapitel=kapitel,
            words=prepared_words,
        )
        return CategoryProcessResult(
            sheet_name=sheet_name,
            added_count=sync_result.added_count,
            duplicate_count=sync_result.duplicate_count,
            priority_updated_count=sync_result.priority_updated_count,
            planned_actions=sync_result.planned_actions,
        )

    def _resolve_sheet_name(self, category: str) -> str:
        category_map = self.config.get("category_to_sheet", {})
        return category_map.get(category, category)
