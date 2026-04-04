import re
import socket
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from utils import normalize_text


@dataclass(slots=True)
class PreparedWord:
    term: str
    definition: str
    priority: int = 0


@dataclass(slots=True)
class SyncResult:
    added_count: int = 0
    duplicate_count: int = 0
    priority_updated_count: int = 0
    planned_actions: list["PlannedAction"] | None = None


@dataclass(slots=True)
class SheetLayout:
    index_column: str = "A"
    term_column: str = "B"
    definition_column: str = "D"
    module_link_column: str = "E"
    priority_column: str = "G"
    data_start_row: int = 2
    teil_size: int = 20
    default_priority: int = 0
    header_fill_color: str = "#3f9744"
    priority_labels: dict[int, str] | None = None

    @classmethod
    def from_config(cls, config: dict[str, Any]) -> "SheetLayout":
        return cls(
            index_column=config.get("index_column", "A"),
            term_column=config.get("term_column", "B"),
            definition_column=config.get("definition_column", "D"),
            module_link_column=config.get("module_link_column", "E"),
            priority_column=config.get("priority_column", "G"),
            data_start_row=int(config.get("data_start_row", 2)),
            teil_size=int(config.get("teil_size", 20)),
            default_priority=int(config.get("default_priority", 0)),
            header_fill_color=config.get("header_fill_color", "#3f9744"),
            priority_labels={
                int(priority): str(label)
                for priority, label in config.get("priority_labels", {}).items()
            }
            or None,
        )


@dataclass(slots=True)
class KapitelHeader:
    kapitel: int
    row_number: int


@dataclass(slots=True)
class KapitelBlock:
    kapitel: int
    teil: int
    header_row: int
    data_start_row: int
    data_end_row: int
    module_label: str


@dataclass(slots=True)
class Slot:
    row_number: int
    item_index: int
    teil: int
    module_label: str


@dataclass(slots=True)
class KapitelState:
    slots: list[Slot]
    next_teil: int
    insert_row: int


@dataclass(slots=True)
class PlannedAction:
    action: str
    row_number: int | None
    teil: int | None
    term: str
    details: str = ""


class GoogleSheetsHandler:
    def __init__(
        self,
        sheet_id: str,
        credentials_file: str,
        sheet_layout: SheetLayout | None = None,
        dry_run: bool = False,
    ) -> None:
        self.sheet_id = sheet_id
        self.credentials_file = credentials_file
        self.sheet_layout = sheet_layout or SheetLayout()
        self.dry_run = dry_run
        self.service = None
        self.sheet_title_to_id: dict[str, int] = {}
        self.sheet_title_to_row_count: dict[str, int] = {}
        self.sheet_title_to_column_count: dict[str, int] = {}
        self.priority_labels = self.sheet_layout.priority_labels or {
            0: "0",
            1: "1",
            2: "2",
            3: "3",
        }

    @classmethod
    def from_config(cls, config: dict, dry_run: bool = False) -> "GoogleSheetsHandler":
        google_config = config.get("google_sheets", {})
        return cls(
            sheet_id=google_config.get("sheet_id", ""),
            credentials_file=google_config.get("credentials_file", ""),
            sheet_layout=SheetLayout.from_config(config.get("sheet_layout", {})),
            dry_run=dry_run,
        )

    def connect(self) -> None:
        if self.service is not None:
            return

        if not self.sheet_id:
            raise ValueError("В config.json не указан google_sheets.sheet_id")
        if not self.credentials_file:
            raise ValueError("В config.json не указан google_sheets.credentials_file")

        credentials_path = Path(self.credentials_file)
        if not credentials_path.exists():
            raise FileNotFoundError(f"Не найден файл service account: {credentials_path}")

        try:
            from google.oauth2.service_account import Credentials
            import googleapiclient.discovery as discovery
            import httplib2
            from google_auth_httplib2 import AuthorizedHttp
        except ImportError as exc:
            raise ImportError(
                "Не установлены зависимости. Установи: pip install google-api-python-client google-auth"
            ) from exc

        original_create_method = discovery.createMethod

        def safe_create_method(method_name, method_desc, root_desc, schema):
            try:
                return original_create_method(method_name, method_desc, root_desc, schema)
            except MemoryError:
                method_name = discovery.fix_method_name(method_name)
                (
                    path_url,
                    http_method,
                    method_id,
                    accept,
                    max_size,
                    media_path_url,
                ) = discovery._fix_up_method_description(method_desc, root_desc, schema)
                parameters = discovery.ResourceMethodParameters(method_desc)

                def method(self, **kwargs):
                    self._validate_credentials()

                    for name in kwargs:
                        if name not in parameters.argmap:
                            raise TypeError(f"Got an unexpected keyword argument {name}")

                    keys = list(kwargs.keys())
                    for name in keys:
                        if kwargs[name] is None:
                            del kwargs[name]

                    for name in parameters.required_params:
                        if name not in kwargs:
                            if name not in discovery._PAGE_TOKEN_NAMES or discovery._findPageTokenName(
                                discovery._methodProperties(method_desc, schema, "response")
                            ):
                                raise TypeError(f'Missing required parameter "{name}"')

                    for name, regex in parameters.pattern_params.items():
                        if name in kwargs:
                            pvalues = [kwargs[name]] if isinstance(kwargs[name], str) else kwargs[name]
                            for pvalue in pvalues:
                                if re.match(regex, pvalue) is None:
                                    raise TypeError(
                                        f'Parameter "{name}" value "{pvalue}" does not match the pattern "{regex}"'
                                    )

                    for name, enums in parameters.enum_params.items():
                        if name in kwargs:
                            if name in parameters.repeated_params and not isinstance(kwargs[name], str):
                                values = kwargs[name]
                            else:
                                values = [kwargs[name]]
                            for value in values:
                                if value not in enums:
                                    raise TypeError(
                                        f'Parameter "{name}" value "{value}" is not an allowed value in "{enums}"'
                                    )

                    actual_query_params = {}
                    actual_path_params = {}
                    for key, value in kwargs.items():
                        to_type = parameters.param_types.get(key, "string")
                        if key in parameters.repeated_params and isinstance(value, list):
                            cast_value = [discovery._cast(x, to_type) for x in value]
                        else:
                            cast_value = discovery._cast(value, to_type)
                        if key in parameters.query_params:
                            actual_query_params[parameters.argmap[key]] = cast_value
                        if key in parameters.path_params:
                            actual_path_params[parameters.argmap[key]] = cast_value

                    body_value = kwargs.get("body")
                    media_filename = kwargs.get("media_body")
                    media_mime_type = kwargs.get("media_mime_type")

                    if self._developerKey:
                        actual_query_params["key"] = self._developerKey

                    model = self._model
                    if method_name.endswith("_media"):
                        model = discovery.MediaModel()
                    elif "response" not in method_desc:
                        model = discovery.RawModel()

                    api_version = method_desc.get("apiVersion")
                    headers = {}
                    headers, params, query, body = model.request(
                        headers, actual_path_params, actual_query_params, body_value, api_version
                    )

                    expanded_url = discovery.uritemplate.expand(path_url, params)
                    url = discovery._urljoin(self._baseUrl, expanded_url + query)

                    resumable = None
                    if media_filename:
                        raise NotImplementedError(
                            "Media upload methods are not supported by the safe discovery fallback."
                        )

                    return self._requestBuilder(
                        self._http,
                        model.response,
                        url,
                        method=http_method,
                        body=body,
                        headers=headers,
                        methodId=method_id,
                        resumable=resumable,
                    )

                method.__doc__ = discovery.DEFAULT_METHOD_DOC
                return method_name, method

        discovery.createMethod = safe_create_method

        credentials = Credentials.from_service_account_file(
            str(credentials_path),
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        previous_socket_timeout = socket.getdefaulttimeout()
        socket.setdefaulttimeout(15)
        http = httplib2.Http(timeout=30)
        authed_http = AuthorizedHttp(credentials, http=http)

        try:
            self.service = discovery.build(
                "sheets",
                "v4",
                http=authed_http,
                cache_discovery=False,
            )
            self.sheet_title_to_id = self._load_sheet_titles()
        except socket.timeout as exc:
            raise TimeoutError(
                "Google Sheets API не ответил за 15 секунд. Проверь интернет, VPN/прокси, firewall и доступ к Google."
            ) from exc
        except TimeoutError as exc:
            raise TimeoutError(
                "Google Sheets API не ответил за 30 секунд. Проверь интернет, доступ к Google API и попробуй снова."
            ) from exc
        except OSError as exc:
            raise ConnectionError(
                f"Не удалось подключиться к Google Sheets API: {exc}"
            ) from exc
        finally:
            socket.setdefaulttimeout(previous_socket_timeout)

    def sync_category_words(
        self,
        sheet_name: str,
        category: str,
        kapitel: int,
        words: list[PreparedWord],
    ) -> SyncResult:
        del category
        self.connect()
        self._ensure_sheet_exists(sheet_name)

        rows = self._get_sheet_values(sheet_name)
        existing_words = self._get_existing_words(rows)
        kapitel_state = self._get_kapitel_state(rows, sheet_name, kapitel)
        result = SyncResult()
        planned_actions: list[PlannedAction] = []

        value_updates: list[dict[str, Any]] = []
        new_blocks: list[tuple[int, list[list[str]]]] = []

        for word in words:
            normalized_term = normalize_text(word.term)
            duplicate_info = existing_words.get(normalized_term)
            if duplicate_info is not None:
                result.duplicate_count += 1
                priority_row, current_priority = duplicate_info
                next_priority = min(max(current_priority + 1, 1), 3)
                if next_priority != current_priority and not self.dry_run:
                    value_updates.append(
                        self._build_cell_update(
                            sheet_name=sheet_name,
                            column=self.sheet_layout.priority_column,
                            row_number=priority_row,
                            value=self._format_priority_value(next_priority),
                        )
                    )
                    result.priority_updated_count += 1
                elif next_priority != current_priority:
                    result.priority_updated_count += 1
                planned_actions.append(
                    PlannedAction(
                        action="priority_update",
                        row_number=priority_row,
                        teil=None,
                        term=word.term,
                        details=(
                            f"{self._format_priority_value(current_priority)}"
                            f" -> {self._format_priority_value(next_priority)}"
                        ),
                    )
                )
                continue

            if kapitel_state.slots:
                slot = kapitel_state.slots.pop(0)
                row_values = self._build_word_row(
                    item_index=slot.item_index,
                    term=word.term,
                    definition=word.definition,
                    module_label=slot.module_label,
                    priority=word.priority,
                )
                if not self.dry_run:
                    value_updates.append(self._build_row_update(sheet_name, slot.row_number, row_values))
                existing_words[normalized_term] = (slot.row_number, word.priority)
                result.added_count += 1
                planned_actions.append(
                    PlannedAction(
                        action="fill_slot",
                        row_number=slot.row_number,
                        teil=slot.teil,
                        term=word.term,
                    )
                )
                continue

            if not new_blocks or self._find_next_empty_block_row(new_blocks[-1][1]) is None:
                new_block_teil = kapitel_state.next_teil + len(new_blocks)
                module_label = self._build_module_label(sheet_name, kapitel, new_block_teil)
                new_block_rows = self._build_empty_block_rows(module_label)
                new_blocks.append((new_block_teil, new_block_rows))
                planned_actions.append(
                    PlannedAction(
                        action="create_teil",
                        row_number=kapitel_state.insert_row + (len(new_blocks) - 1) * len(new_block_rows),
                        teil=new_block_teil,
                        term="",
                        details=f"insert before row {kapitel_state.insert_row + (len(new_blocks) - 1) * len(new_block_rows)}",
                    )
                )

            new_block_teil, new_block_rows = new_blocks[-1]
            next_slot_idx = self._find_next_empty_block_row(new_block_rows)
            if next_slot_idx is None:
                raise ValueError("Не удалось найти пустой слот в новом блоке.")
            row_values = self._build_word_row(
                item_index=next_slot_idx,
                term=word.term,
                definition=word.definition,
                module_label=self._build_module_label(sheet_name, kapitel, new_block_teil),
                priority=word.priority,
            )
            new_block_rows[next_slot_idx] = row_values
            result.added_count += 1
            planned_actions.append(
                PlannedAction(
                    action="fill_new_teil",
                    row_number=kapitel_state.insert_row + (len(new_blocks) - 1) * len(new_block_rows) + next_slot_idx,
                    teil=new_block_teil,
                    term=word.term,
                )
            )

        if not self.dry_run:
            self._apply_mutations(
                sheet_name=sheet_name,
                kapitel=kapitel,
                insert_row=kapitel_state.insert_row,
                new_blocks=new_blocks,
                value_updates=value_updates,
            )

        result.planned_actions = planned_actions
        return result

    def _get_kapitel_state(self, rows: list[list[str]], sheet_name: str, kapitel: int) -> KapitelState:
        kapitel_headers = self._find_kapitel_headers(rows)
        is_sprechen_sheet = self._is_redemittel_sprechen_sheet(sheet_name) or self._is_sprechen_sheet_by_structure(rows)
        teil_blocks = self._find_teil_blocks(rows, sheet_name, kapitel, force_sprechen=is_sprechen_sheet)

        if is_sprechen_sheet:
            blocks = [block for block in teil_blocks if block.kapitel == kapitel]
            blocks.sort(key=lambda item: item.header_row)
            slots: list[Slot] = []
            for block in blocks:
                slots.extend(self._collect_free_slots(rows, block))

            next_teil = max((block.teil for block in blocks), default=1) + 1
            first_kapitel_header_row = min(
                (header.row_number for header in kapitel_headers),
                default=len(rows) + 1,
            )
            insert_row = max((block.data_end_row for block in blocks), default=0) + 1
            insert_row = min(insert_row, first_kapitel_header_row)
            return KapitelState(slots=slots, next_teil=next_teil, insert_row=insert_row)

        self._validate_kapitel_structure(kapitel_headers, teil_blocks, kapitel)

        kapitel_header = next((header for header in kapitel_headers if header.kapitel == kapitel), None)

        if kapitel_header is None:
            blocks = [block for block in teil_blocks if block.kapitel == kapitel]
            next_kapitel_header_row = len(rows) + 1
            section_anchor_row = min((block.header_row for block in blocks), default=len(rows) + 1)
        else:
            next_kapitel_header_row = next(
                (header.row_number for header in kapitel_headers if header.row_number > kapitel_header.row_number),
                len(rows) + 1,
            )
            section_anchor_row = kapitel_header.row_number
            blocks = [
                block
                for block in teil_blocks
                if block.kapitel == kapitel and kapitel_header.row_number < block.header_row < next_kapitel_header_row
            ]

        blocks.sort(key=lambda item: (item.teil, item.header_row))

        slots: list[Slot] = []
        for block in blocks:
            slots.extend(self._collect_free_slots(rows, block))

        next_teil = max((block.teil for block in blocks), default=0) + 1
        insert_row = max((block.data_end_row for block in blocks), default=section_anchor_row) + 1
        insert_row = min(insert_row, next_kapitel_header_row)

        return KapitelState(slots=slots, next_teil=next_teil, insert_row=insert_row)

    @staticmethod
    def _validate_kapitel_structure(
        kapitel_headers: list[KapitelHeader],
        teil_blocks: list[KapitelBlock],
        kapitel: int,
    ) -> None:
        matching_headers = [header for header in kapitel_headers if header.kapitel == kapitel]
        if len(matching_headers) > 1:
            rows = ", ".join(str(header.row_number) for header in matching_headers)
            raise ValueError(
                f"Найдены дубли заголовка 'Kapitel {kapitel} :' в строках: {rows}. "
                "Скрипт остановлен, чтобы не повредить лист."
            )

        # Некоторые листы уже содержат повторяющиеся блоки Teil внутри одного Kapitel.
        # Вместо остановки используем все найденные блоки как отдельные секции,
        # чтобы сначала заполнять их пустые слоты сверху вниз.

    def _collect_free_slots(self, rows: list[list[str]], block: KapitelBlock) -> list[Slot]:
        index_col = self._column_to_index(self.sheet_layout.index_column)
        term_col = self._column_to_index(self.sheet_layout.term_column)
        definition_col = self._column_to_index(self.sheet_layout.definition_column)
        module_col = self._column_to_index(self.sheet_layout.module_link_column)
        slots: list[Slot] = []

        for expected_index, row_number in enumerate(
            range(block.data_start_row, block.data_end_row + 1),
            start=1,
        ):
            row = rows[row_number - 1] if row_number - 1 < len(rows) else []
            term_value = self._get_row_value(row, term_col)
            definition_value = self._get_row_value(row, definition_col)
            module_value = self._get_row_value(row, module_col)
            item_index = self._parse_int(self._get_row_value(row, index_col), expected_index)

            if self._is_slot_occupied(term_value, definition_value, module_value, block):
                continue

            slots.append(
                Slot(
                    row_number=row_number,
                    item_index=item_index,
                    teil=block.teil,
                    module_label=block.module_label,
                )
            )

        return slots

    @staticmethod
    def _is_slot_occupied(
        term_value: str,
        definition_value: str,
        module_value: str,
        block: KapitelBlock,
    ) -> bool:
        if term_value or definition_value:
            return True
        normalized_module = normalize_text(module_value)
        return normalized_module not in ("", normalize_text(block.module_label))

    def _apply_mutations(
        self,
        sheet_name: str,
        kapitel: int,
        insert_row: int,
        new_blocks: list[tuple[int, list[list[str]]]],
        value_updates: list[dict[str, Any]],
    ) -> None:
        sheet_id = self.sheet_title_to_id[sheet_name]
        self._ensure_sheet_has_columns(sheet_name, sheet_id, 7)

        if new_blocks:
            requests: list[dict[str, Any]] = []
            inserted_row_count = 0
            for teil, new_block_rows in new_blocks:
                block_insert_row = insert_row + inserted_row_count
                requests.extend(
                    self._build_new_block_requests(
                        sheet_name=sheet_name,
                        sheet_id=sheet_id,
                        insert_row=block_insert_row,
                        kapitel=kapitel,
                        teil=teil,
                        row_count=len(new_block_rows),
                    )
                )
                inserted_row_count += len(new_block_rows)
            (
                self.service.spreadsheets()
                .batchUpdate(
                    spreadsheetId=self.sheet_id,
                    body={"requests": requests},
                )
                .execute()
            )
            self.sheet_title_to_row_count[sheet_name] = (
                self.sheet_title_to_row_count.get(sheet_name, 0) + inserted_row_count
            )

            inserted_row_count = 0
            for _, new_block_rows in new_blocks:
                for offset, row_values in enumerate(new_block_rows[1:], start=1):
                    value_updates.append(
                        self._build_row_update(sheet_name, insert_row + inserted_row_count + offset, row_values)
                    )
                inserted_row_count += len(new_block_rows)

        if value_updates:
            (
                self.service.spreadsheets()
                .values()
                .batchUpdate(
                    spreadsheetId=self.sheet_id,
                    body={
                        "valueInputOption": "USER_ENTERED",
                        "data": value_updates,
                    },
                )
                .execute()
            )

    def _build_new_block_requests(
        self,
        sheet_name: str,
        sheet_id: int,
        insert_row: int,
        kapitel: int,
        teil: int,
        row_count: int,
    ) -> list[dict[str, Any]]:
        start_index = insert_row - 1
        header_text = self._build_block_header_text(sheet_name, kapitel, teil)
        color = self._hex_to_google_color(self._get_header_fill_color(sheet_name))
        return [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "ROWS",
                        "startIndex": start_index,
                        "endIndex": start_index + row_count,
                    },
                    "inheritFromBefore": True,
                }
            },
            {
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_index,
                        "endRowIndex": start_index + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 7,
                    },
                    "mergeType": "MERGE_ALL",
                }
            },
            {
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_index,
                        "endRowIndex": start_index + 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 7,
                    },
                    "rows": [
                        {
                            "values": [
                                {
                                    "userEnteredValue": {"stringValue": header_text},
                                    "userEnteredFormat": {
                                        "backgroundColor": color,
                                        "horizontalAlignment": "CENTER",
                                        "textFormat": {
                                            "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                                            "bold": True,
                                        },
                                    },
                                }
                            ]
                        }
                    ],
                    "fields": "userEnteredValue,userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_index + 1,
                        "endRowIndex": start_index + row_count,
                        "startColumnIndex": 0,
                        "endColumnIndex": 7,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {"red": 1, "green": 1, "blue": 1},
                            "horizontalAlignment": "LEFT",
                            "textFormat": {
                                "foregroundColor": {"red": 0, "green": 0, "blue": 0},
                                "bold": False,
                            },
                        }
                    },
                    "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,textFormat)",
                }
            },
        ]

    def _build_empty_block_rows(self, module_label: str) -> list[list[str]]:
        rows: list[list[str]] = [[""] * 7]
        for item_index in range(1, self.sheet_layout.teil_size + 1):
            rows.append(
                self._build_word_row(
                    item_index=item_index,
                    term="",
                    definition="",
                    module_label=module_label,
                    priority=self.sheet_layout.default_priority,
                )
            )
        return rows

    @staticmethod
    def _count_filled_rows(new_block_rows: list[list[str]]) -> int:
        term_col_index = 1
        filled_count = 0
        for row_values in new_block_rows[1:]:
            if row_values[term_col_index]:
                filled_count += 1
        return filled_count + 1

    @staticmethod
    def _find_next_empty_block_row(new_block_rows: list[list[str]]) -> int | None:
        term_col_index = 1
        for row_index, row_values in enumerate(new_block_rows[1:], start=1):
            if not row_values[term_col_index]:
                return row_index
        return None

    def _get_existing_words(self, rows: list[list[str]]) -> dict[str, tuple[int, int]]:
        term_col = self._column_to_index(self.sheet_layout.term_column)
        priority_col = self._column_to_index(self.sheet_layout.priority_column)
        existing_words: dict[str, tuple[int, int]] = {}

        for row_number, row in enumerate(rows, start=1):
            term = self._get_row_value(row, term_col)
            if not term or self._looks_like_header(term):
                continue
            priority = self._parse_int(self._get_row_value(row, priority_col), self.sheet_layout.default_priority)
            priority = self._parse_priority_value(self._get_row_value(row, priority_col), priority)
            existing_words[normalize_text(term)] = (row_number, priority)

        return existing_words

    def _find_kapitel_headers(self, rows: list[list[str]]) -> list[KapitelHeader]:
        headers: list[KapitelHeader] = []
        for row_number, row in enumerate(rows, start=1):
            text = self._row_text(row)
            normalized = normalize_text(text)
            if re.search(r"(?:\u0447\u0430\u0441\u0442\u044c|teil)\s+\d+", normalized):
                continue
            match = re.fullmatch(r"kapitel\s+(\d+)\s*:\s*.*", normalized)
            if match:
                headers.append(KapitelHeader(kapitel=int(match.group(1)), row_number=row_number))
        return headers

    def _find_teil_blocks(
        self,
        rows: list[list[str]],
        sheet_name: str,
        kapitel: int,
        force_sprechen: bool = False,
    ) -> list[KapitelBlock]:
        blocks: list[KapitelBlock] = []
        if force_sprechen or self._is_redemittel_sprechen_sheet(sheet_name):
            for row_number, row in enumerate(rows, start=1):
                row_text = self._row_text(row)
                teil = self._parse_sprechen_teil_header(row_text)
                if teil is None:
                    continue
                blocks.append(
                    KapitelBlock(
                        kapitel=kapitel,
                        teil=teil,
                        header_row=row_number,
                        data_start_row=row_number + 1,
                        data_end_row=row_number + self.sheet_layout.teil_size,
                        module_label=self._build_module_label(sheet_name, kapitel, teil),
                    )
                )
            return blocks

        current_kapitel: int | None = None
        for row_number, row in enumerate(rows, start=1):
            row_text = self._row_text(row)
            kapitel_header = self._parse_kapitel_header(row_text)
            if kapitel_header is not None:
                current_kapitel = kapitel_header

            parsed = self._parse_teil_header(row_text)
            if parsed is None:
                continue
            kapitel, teil = parsed
            if kapitel == 0:
                if current_kapitel is None:
                    continue
                kapitel = current_kapitel
            blocks.append(
                KapitelBlock(
                    kapitel=kapitel,
                    teil=teil,
                    header_row=row_number,
                    data_start_row=row_number + 1,
                    data_end_row=row_number + self.sheet_layout.teil_size,
                    module_label=self._build_module_label(sheet_name, kapitel, teil),
                )
            )
        return blocks

    @staticmethod
    def _row_text(row: list[str]) -> str:
        return " ".join(cell.strip() for cell in row if cell and cell.strip())

    @staticmethod
    def _parse_kapitel_header(value: str) -> int | None:
        text = normalize_text(value)
        if re.search(r"(?:\u0447\u0430\u0441\u0442\u044c|teil)\s+\d+", text):
            return None
        match = re.fullmatch(r"kapitel\s+(\d+)\s*:\s*.*", text)
        if not match:
            return None
        return int(match.group(1))

    @staticmethod
    def _parse_teil_header(value: str) -> tuple[int, int] | None:
        text = normalize_text(value)
        match = re.search(r"kapitel\s+(\d+)\s*:\s*(?:\u0447\u0430\u0441\u0442\u044c|teil)\s+(\d+)", text)
        if match:
            return int(match.group(1)), int(match.group(2))
        match = re.fullmatch(r"(?:\u0447\u0430\u0441\u0442\u044c|teil)\s+(\d+)", text)
        if not match:
            return None
        return 0, int(match.group(1))

    @staticmethod
    def _build_cell_update(
        sheet_name: str,
        column: str,
        row_number: int,
        value: str,
    ) -> dict[str, Any]:
        return {"range": f"'{sheet_name}'!{column}{row_number}", "values": [[value]]}

    @staticmethod
    def _build_row_update(sheet_name: str, row_number: int, row_values: list[str]) -> dict[str, Any]:
        return {"range": f"'{sheet_name}'!A{row_number}:G{row_number}", "values": [row_values]}

    def _build_word_row(
        self,
        item_index: int,
        term: str,
        definition: str,
        module_label: str,
        priority: int,
    ) -> list[str]:
        row = [""] * 7
        row[self._column_to_index(self.sheet_layout.index_column) - 1] = str(item_index)
        row[self._column_to_index(self.sheet_layout.term_column) - 1] = term
        row[self._column_to_index(self.sheet_layout.definition_column) - 1] = definition
        row[self._column_to_index(self.sheet_layout.module_link_column) - 1] = module_label
        row[self._column_to_index(self.sheet_layout.priority_column) - 1] = self._format_priority_value(priority)
        return row

    def _format_priority_value(self, priority: int) -> str:
        return self.priority_labels.get(priority, str(priority))

    def _parse_priority_value(self, value: str, fallback: int) -> int:
        normalized_value = value.strip()
        if not normalized_value:
            return fallback

        for priority, label in self.priority_labels.items():
            if normalized_value == label:
                return priority

        match = re.match(r"\s*(\d+)", normalized_value)
        if match:
            return min(int(match.group(1)), 3)

        return fallback

    def _load_sheet_titles(self) -> dict[str, int]:
        response = (
            self.service.spreadsheets()
            .get(
                spreadsheetId=self.sheet_id,
                fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
            )
            .execute()
        )
        result: dict[str, int] = {}
        for sheet in response.get("sheets", []):
            properties = sheet.get("properties", {})
            title = properties.get("title", "")
            result[title] = properties.get("sheetId", 0)
            self.sheet_title_to_row_count[title] = properties.get("gridProperties", {}).get("rowCount", 0)
            self.sheet_title_to_column_count[title] = properties.get("gridProperties", {}).get("columnCount", 0)
        return result

    def _ensure_sheet_exists(self, sheet_name: str) -> None:
        if sheet_name not in self.sheet_title_to_id:
            available_titles = ", ".join(self.sheet_title_to_id.keys())
            raise ValueError(f"Лист '{sheet_name}' не найден в таблице. Доступные листы: {available_titles}")

    def _get_sheet_values(self, sheet_name: str) -> list[list[str]]:
        response = (
            self.service.spreadsheets()
            .values()
            .batchGet(
                spreadsheetId=self.sheet_id,
                ranges=[
                    f"'{sheet_name}'!A:B",
                    f"'{sheet_name}'!E:G",
                ],
            )
            .execute()
        )
        value_ranges = response.get("valueRanges", [])
        left_values = value_ranges[0].get("values", []) if len(value_ranges) > 0 else []
        right_values = value_ranges[1].get("values", []) if len(value_ranges) > 1 else []

        row_count = max(len(left_values), len(right_values))
        merged_rows: list[list[str]] = []
        for row_index in range(row_count):
            left_row = left_values[row_index] if row_index < len(left_values) else []
            right_row = right_values[row_index] if row_index < len(right_values) else []

            row = [""] * 7
            for index, value in enumerate(left_row[:2]):
                row[index] = value
            for index, value in enumerate(right_row[:3], start=4):
                row[index] = value
            merged_rows.append(row)

        return merged_rows

    def _ensure_sheet_has_columns(self, sheet_name: str, sheet_id: int, required_column_count: int) -> None:
        current_column_count = self.sheet_title_to_column_count.get(sheet_name, 0)
        if required_column_count <= current_column_count:
            return
        (
            self.service.spreadsheets()
            .batchUpdate(
                spreadsheetId=self.sheet_id,
                body={
                    "requests": [
                        {
                            "appendDimension": {
                                "sheetId": sheet_id,
                                "dimension": "COLUMNS",
                                "length": required_column_count - current_column_count,
                            }
                        }
                    ]
                },
            )
            .execute()
        )
        self.sheet_title_to_column_count[sheet_name] = required_column_count

    @staticmethod
    def _get_row_value(row: list[str], column_index: int) -> str:
        if column_index - 1 >= len(row):
            return ""
        return row[column_index - 1].strip()

    @staticmethod
    def _parse_int(value: str, default: int) -> int:
        try:
            return int(str(value).strip())
        except (TypeError, ValueError):
            return default

    @staticmethod
    def _column_to_index(column_name: str) -> int:
        result = 0
        for character in column_name.upper():
            result = result * 26 + ord(character) - ord("A") + 1
        return result

    @staticmethod
    def _hex_to_google_color(color: str) -> dict[str, float]:
        color = color.lstrip("#")
        return {
            "red": int(color[0:2], 16) / 255,
            "green": int(color[2:4], 16) / 255,
            "blue": int(color[4:6], 16) / 255,
        }

    @staticmethod
    def _looks_like_header(value: str) -> bool:
        normalized = normalize_text(value)
        return normalized.startswith("fokus deutsch") or normalized.startswith("kapitel ")

    @staticmethod
    def _is_redemittel_sprechen_sheet(sheet_name: str) -> bool:
        normalized = normalize_text(sheet_name)
        return normalized == normalize_text("🗣️ REDEMITTEL Sprechen") or "redemittel sprechen" in normalized

    def _is_sprechen_sheet_by_structure(self, rows: list[list[str]]) -> bool:
        sprechen_header_count = 0
        kapitel_header_count = 0

        for row in rows:
            row_text = self._row_text(row)
            if not row_text:
                continue
            if self._parse_kapitel_header(row_text) is not None:
                kapitel_header_count += 1
            if self._parse_sprechen_teil_header(row_text) is not None:
                sprechen_header_count += 1

        return sprechen_header_count >= 2 and kapitel_header_count == 0

    def _get_header_fill_color(self, sheet_name: str) -> str:
        if self._is_redemittel_sprechen_sheet(sheet_name):
            return "#5f78bd"
        return self.sheet_layout.header_fill_color

    def _build_block_header_text(self, sheet_name: str, kapitel: int, teil: int) -> str:
        if self._is_redemittel_sprechen_sheet(sheet_name):
            return f"Sprechen B2 - Teil {teil}"
        return f"Kapitel {kapitel} : часть {teil}"

    def _build_module_label(self, sheet_name: str, kapitel: int, teil: int) -> str:
        if self._is_redemittel_sprechen_sheet(sheet_name):
            return f"Sprechen B2 - Teil {teil}"
        return f"Kapitel {kapitel} - \u0447\u0430\u0441\u0442\u044c {teil}"

    @staticmethod
    def _parse_sprechen_teil_header(value: str) -> int | None:
        text = normalize_text(value)
        match = re.fullmatch(r"sprechen\s*b2\s*-\s*teil\s+(\d+)", text)
        if not match:
            return None
        return int(match.group(1))

    @staticmethod
    def _parse_module_label(value: str) -> tuple[int, int] | None:
        text = normalize_text(value)
        match = re.fullmatch(r"kapitel\s+(\d+)\s*-\s*(?:\u0447\u0430\u0441\u0442\u044c|teil)\s+(\d+)", text)
        if not match:
            return None
        return int(match.group(1)), int(match.group(2))


