"""Streamlit-приложение для поиска занятий преподавателя в Excel-расписании."""

from __future__ import annotations

import logging
import math
import re
from collections import Counter
from collections.abc import MutableMapping, Sequence
from dataclasses import dataclass
from datetime import datetime, time
from hashlib import sha256
from io import BytesIO
from numbers import Real
from typing import Any, BinaryIO, Final

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


LOGGER = logging.getLogger(__name__)

GROUP_PATTERN: Final[re.Pattern[str]] = re.compile(r"[А-Я]{4}-\d{2}-\d{2}")
WHITESPACE_PATTERN: Final[re.Pattern[str]] = re.compile(r"\s+")
PAIR_NUMBER_PATTERN: Final[re.Pattern[str]] = re.compile(r"\d+(?:[.,]\d+)?")
TIME_TEXT_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"^(?P<hours>\d{1,2})\s*[:.\-]\s*(?P<minutes>\d{2})"
    r"(?:\s*[:.\-]\s*(?P<seconds>\d{2}))?$"
)
WORKLOAD_SUBJECT_SUFFIX_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"\s*\((?=[^)]*(?:з\.?\s*е\.?|\bч\.?|УП\s*№|Д\s*№))[^)]*\)\s*$",
    re.IGNORECASE,
)
NAME_TOKEN_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"[а-яёa-z0-9+#]+",
    re.IGNORECASE,
)

HEADER_SEARCH_ROWS: Final[int] = 10
COMMON_COLUMNS_COUNT: Final[int] = 5
FIRST_GROUP_COLUMN: Final[int] = 5
GROUP_BLOCK_SIZE: Final[int] = 5
SESSION_CACHE_KEY: Final[str] = "_schedule_excel_cache"
MAX_CACHED_FILES: Final[int] = 20
APP_VERSION: Final[str] = "1.4.0"
APP_CHANGELOG: Final[tuple[tuple[str, str, tuple[str, ...]], ...]] = (
    (
        "1.4.0",
        "28.08.2026",
        (
            "Добавлена необязательная сверка с файлом учебной нагрузки.",
            "Добавлен выбор осеннего или весеннего семестра.",
            "Ошибки нагрузки и возможные места для ФИО показаны отдельно.",
        ),
    ),
    (
        "1.3.0",
        "28.08.2026",
        (
            "Добавлено постоянное отображение версии приложения.",
            "Добавлен кликабельный журнал изменений в нижней части страницы.",
        ),
    ),
    (
        "1.2.0",
        "28.08.2026",
        (
            "Общие лекции для нескольких групп объединяются без предупреждения.",
            "Разные занятия одного временного слота выводятся в одной строке.",
            "Накладки выделяются в таблице сайта и XLSX-выгрузке.",
        ),
    ),
    (
        "1.1.0",
        "27.08.2026",
        (
            "Добавлена пакетная загрузка файлов разных курсов.",
            "Добавлен поиск сразу нескольких преподавателей.",
            "Добавлена цветная XLSX-выгрузка расписания.",
        ),
    ),
    (
        "1.0.0",
        "27.08.2026",
        ("Первая версия парсера грязных Excel-файлов расписания.",),
    ),
)

QUERY_COLUMN: Final[str] = "Искомый преподаватель"
TEACHER_COLUMN: Final[str] = "Преподаватель"
SOURCE_COLUMN: Final[str] = "Файл-источник"
CONFLICT_COLUMN: Final[str] = "Накладка"
WORKLOAD_CLASS_TYPES: Final[frozenset[str]] = frozenset({"ЛК", "ПР"})

OUTPUT_COLUMNS: Final[list[str]] = [
    "День недели",
    "Пара",
    "Время",
    "Неделя",
    "Группа",
    "Дисциплина",
    "Вид занятий",
    "Аудитория",
]

INTERNAL_COLUMNS: Final[list[str]] = [
    QUERY_COLUMN,
    TEACHER_COLUMN,
    *OUTPUT_COLUMNS,
    SOURCE_COLUMN,
]

ALL_SCHEDULE_COLUMNS: Final[list[str]] = [
    TEACHER_COLUMN,
    *OUTPUT_COLUMNS,
    SOURCE_COLUMN,
]

WORKLOAD_COLUMNS: Final[list[str]] = [
    "Преподаватель нагрузки",
    "Дисциплина",
    "Вид занятий",
    "Группа",
    "Семестр",
    "Строка нагрузки",
]

WORKLOAD_ISSUE_COLUMNS: Final[list[str]] = [
    "Статус",
    "Проблема",
    "Преподаватели по нагрузке",
    "Дисциплина",
    "Вид занятий",
    "Группа",
    "Семестр",
    "ФИО в расписании",
    "Предлагаемое ФИО",
    "Возможная накладка",
    "Возможное место в расписании",
]

DISPLAY_COLUMNS: Final[list[str]] = [
    QUERY_COLUMN,
    TEACHER_COLUMN,
    "День недели",
    "Пара",
    "Время",
    "Неделя",
    CONFLICT_COLUMN,
    "Группа",
    "Дисциплина",
    "Вид занятий",
    "Аудитория",
    SOURCE_COLUMN,
]

EXPORT_SUBHEADERS: Final[tuple[str, str, str, str]] = (
    "Дисциплина",
    "Вид",
    "Группа",
    "Аудитория",
)

DAY_EXPORT_CONFIG: Final[tuple[tuple[str, str, str, str], ...]] = (
    ("ПН", "FFFF54", "FBE6A2", "F9DA78"),
    ("ВТ", "75FB4C", "BCD6AC", "9DC284"),
    ("СР", "5884E1", "AAC1F0", "789DE5"),
    ("ЧТ", "F6B26B", "FCE5CD", "F9CB9C"),
    ("ПТ", "B57EDC", "E4D3F3", "CAB2E4"),
    ("СБ", "A6A6A6", "F2F2F2", "D9D9D9"),
)

TEACHER_QUERY_SEPARATOR_PATTERN: Final[re.Pattern[str]] = re.compile(r"[,;\n]+")
TEACHER_NAME_SEPARATOR_PATTERN: Final[re.Pattern[str]] = re.compile(r"[,;]+")

DAY_ORDER: Final[dict[str, int]] = {
    "понедельник": 0,
    "пн": 0,
    "вторник": 1,
    "вт": 1,
    "среда": 2,
    "ср": 2,
    "четверг": 3,
    "чт": 3,
    "пятница": 4,
    "пт": 4,
    "суббота": 5,
    "сб": 5,
}


class ScheduleError(Exception):
    """Базовая ожидаемая ошибка обработки расписания."""


class ScheduleReadError(ScheduleError):
    """Excel-файл не удалось прочитать."""


class ScheduleFormatError(ScheduleError):
    """Структура Excel-файла не похожа на ожидаемое расписание."""


class ScheduleExportError(ScheduleError):
    """Не удалось сформировать Excel-файл с итоговым расписанием."""


class WorkloadError(ScheduleError):
    """Базовая ожидаемая ошибка файла учебной нагрузки."""


class WorkloadFormatError(WorkloadError):
    """Структура файла нагрузки не содержит обязательных столбцов."""


@dataclass(frozen=True, slots=True)
class BatchFileError:
    """Ошибка одного файла внутри пакетной обработки."""

    filename: str
    message: str


@dataclass(frozen=True, slots=True)
class BatchParseResult:
    """Результат отказоустойчивой обработки набора Excel-файлов."""

    records: pd.DataFrame
    all_records: pd.DataFrame
    processed_files: tuple[str, ...]
    duplicate_files: tuple[str, ...]
    errors: tuple[BatchFileError, ...]


@dataclass(frozen=True, slots=True)
class WorkloadAuditResult:
    """Результат сверки учебной нагрузки с расписанием."""

    checked_assignments: int
    matched_assignments: int
    errors: pd.DataFrame
    suggestions: pd.DataFrame


def _is_missing(value: Any) -> bool:
    """Безопасно проверяет скалярное значение ячейки на пропуск."""

    if value is None:
        return True

    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        # Ячейка должна быть скаляром, но грязный источник не должен ломать парсер.
        return False


def clean_cell(value: Any) -> str:
    """Преобразует значение Excel-ячейки в чистую однострочную строку."""

    if _is_missing(value):
        return ""

    # Из-за NaN pandas часто превращает целые значения Excel в float (101 -> 101.0).
    if isinstance(value, Real) and not isinstance(value, bool):
        numeric_value = float(value)
        if math.isfinite(numeric_value) and numeric_value.is_integer():
            return str(int(numeric_value))

    text = str(value).replace("\r", " ").replace("\n", " ").replace("\xa0", " ")
    return WHITESPACE_PATTERN.sub(" ", text).strip()


def _normalize_search_text(value: Any) -> str:
    """Нормализует регистр и русские е/ё только для сравнения строк."""

    return clean_cell(value).casefold().replace("ё", "е")


def _normalize_discipline(value: Any) -> str:
    """Нормализует название дисциплины для сверки разных Excel-источников."""

    text = WORKLOAD_SUBJECT_SUFFIX_PATTERN.sub("", clean_cell(value))
    normalized = _normalize_search_text(text)
    return " ".join(NAME_TOKEN_PATTERN.findall(normalized))


def _display_workload_discipline(value: Any) -> str:
    """Убирает из названия нагрузки часы, зачётные единицы и номер плана."""

    return clean_cell(WORKLOAD_SUBJECT_SUFFIX_PATTERN.sub("", clean_cell(value)))


def _normalize_lesson_type(value: Any) -> str:
    """Приводит обозначения лекций и практик к ЛК/ПР."""

    normalized = _normalize_search_text(value)
    if normalized.startswith("лк") or normalized.startswith("лек"):
        return "ЛК"
    if normalized.startswith("пр") or normalized.startswith("практ"):
        return "ПР"
    return clean_cell(value).upper()


def _teacher_surname_key(value: Any) -> str:
    """Возвращает нормализованную фамилию преподавателя."""

    tokens = NAME_TOKEN_PATTERN.findall(_normalize_search_text(value))
    return tokens[0] if tokens else ""


def _extract_groups(*values: Any) -> list[str]:
    """Извлекает названия групп из одного или нескольких грязных полей."""

    groups: list[str] = []
    for value in values:
        for group in GROUP_PATTERN.findall(clean_cell(value).upper()):
            if group not in groups:
                groups.append(group)
    return groups


def _parse_semester_number(value: Any) -> int | None:
    """Безопасно читает номер семестра из числа или строки."""

    if _is_missing(value):
        return None
    if isinstance(value, Real) and not isinstance(value, bool):
        numeric_value = float(value)
        if math.isfinite(numeric_value) and numeric_value.is_integer():
            return int(numeric_value)
    match = re.search(r"\d+", clean_cell(value))
    return int(match.group(0)) if match is not None else None


def parse_teacher_queries(value: str | Sequence[str]) -> list[str]:
    """Разбирает фамилии из строк, запятых и точек с запятой без дублей."""

    if isinstance(value, str):
        raw_queries = TEACHER_QUERY_SEPARATOR_PATTERN.split(value)
    else:
        raw_queries = list(value)

    queries: list[str] = []
    seen: set[str] = set()
    for raw_query in raw_queries:
        query = clean_cell(raw_query)
        normalized_query = _normalize_search_text(query)
        if not normalized_query or normalized_query in seen:
            continue
        seen.add(normalized_query)
        queries.append(query)

    return queries


def _format_pair(value: Any) -> str:
    """Убирает технический суффикс .0 у целочисленных номеров пар."""

    if _is_missing(value):
        return ""

    if isinstance(value, Real) and not isinstance(value, bool):
        numeric_value = float(value)
        if math.isfinite(numeric_value) and numeric_value.is_integer():
            return str(int(numeric_value))

    return clean_cell(value)


def _format_time(value: Any) -> str:
    """Нормализует время из строкового, datetime/time или Excel-числового вида."""

    if _is_missing(value):
        return ""

    if isinstance(value, (pd.Timestamp, datetime, time)):
        return value.strftime("%H:%M")

    # В старых .xls время иногда возвращается долей суток вместо datetime.time.
    if isinstance(value, Real) and not isinstance(value, bool):
        numeric_value = float(value)
        if math.isfinite(numeric_value) and 0 <= numeric_value < 1:
            total_minutes = round(numeric_value * 24 * 60) % (24 * 60)
            hours, minutes = divmod(total_minutes, 60)
            return f"{hours:02d}:{minutes:02d}"

    text_value = clean_cell(value)
    time_match = TIME_TEXT_PATTERN.fullmatch(text_value)
    if time_match is not None and time_match.group("seconds") in (None, "00"):
        return f"{int(time_match.group('hours')):02d}:{time_match.group('minutes')}"

    return text_value


def _build_time_range(start_value: Any, end_value: Any) -> str:
    """Собирает читаемый временной интервал без лишнего разделителя."""

    start = _format_time(start_value)
    end = _format_time(end_value)

    if start and end:
        return f"{start} – {end}"
    return start or end


def _is_time_value(value: Any) -> bool:
    """Проверяет, что значение похоже на время начала или окончания занятия."""

    if _is_missing(value):
        return False
    if isinstance(value, (pd.Timestamp, datetime, time)):
        return True
    if isinstance(value, Real) and not isinstance(value, bool):
        numeric_value = float(value)
        return math.isfinite(numeric_value) and 0 <= numeric_value < 1
    return TIME_TEXT_PATTERN.fullmatch(clean_cell(value)) is not None


def read_excel_file(file: bytes | bytearray | BinaryIO) -> pd.DataFrame:
    """Читает первый лист Excel без предположений о строке заголовков."""

    excel_file: BinaryIO
    if isinstance(file, (bytes, bytearray)):
        excel_file = BytesIO(file)
    else:
        excel_file = file

    try:
        excel_file.seek(0)
        dataframe = pd.read_excel(excel_file, header=None)
    except ImportError as exc:
        raise ScheduleReadError(
            "Не найден модуль для чтения Excel. Установите openpyxl для .xlsx "
            "и xlrd для .xls."
        ) from exc
    except Exception as exc:
        LOGGER.warning("Не удалось прочитать Excel-файл: %s", exc)
        raise ScheduleReadError(
            "Не удалось прочитать файл. Проверьте, что это неповреждённый "
            "Excel-файл формата .xls или .xlsx."
        ) from exc

    if dataframe.empty or dataframe.shape[1] == 0:
        raise ScheduleFormatError("Загруженный Excel-файл не содержит данных.")

    return dataframe


def read_excel_for_session(
    file_content: bytes,
    session_state: MutableMapping[str, Any],
) -> pd.DataFrame:
    """Читает файл один раз на сессию и переиспользует результат при rerun."""

    file_digest = sha256(file_content).hexdigest()
    cached_value = session_state.get(SESSION_CACHE_KEY)

    cache: dict[str, pd.DataFrame] = {}
    if isinstance(cached_value, dict):
        # Миграция кэша из старой версии, где хранился только один файл.
        old_digest = cached_value.get("digest")
        old_dataframe = cached_value.get("dataframe")
        if isinstance(old_digest, str) and isinstance(old_dataframe, pd.DataFrame):
            cache[old_digest] = old_dataframe
        else:
            cache = {
                digest: dataframe
                for digest, dataframe in cached_value.items()
                if isinstance(digest, str) and isinstance(dataframe, pd.DataFrame)
            }

    cached_dataframe = cache.get(file_digest)
    if cached_dataframe is not None:
        # Перемещаем недавно использованный файл в конец insertion-order кэша.
        cache.pop(file_digest)
        cache[file_digest] = cached_dataframe
        session_state[SESSION_CACHE_KEY] = cache
        return cached_dataframe

    dataframe = read_excel_file(file_content)
    cache[file_digest] = dataframe
    while len(cache) > MAX_CACHED_FILES:
        cache.pop(next(iter(cache)))
    session_state[SESSION_CACHE_KEY] = cache
    return dataframe


def find_group_row(df: pd.DataFrame) -> int:
    """Находит позицию строки с названиями групп среди первых десяти строк."""

    rows_to_scan = min(HEADER_SEARCH_ROWS, len(df.index))
    for row_idx in range(rows_to_scan):
        for value in df.iloc[row_idx].tolist():
            if GROUP_PATTERN.search(clean_cell(value)):
                return row_idx

    raise ScheduleFormatError(
        "Не найдена строка с названиями групп в первых 10 строках файла."
    )


def _prepare_schedule_rows(df: pd.DataFrame, group_row_idx: int) -> pd.DataFrame:
    """Выделяет строки расписания и распаковывает вертикальные объединения."""

    if df.shape[1] < COMMON_COLUMNS_COUNT:
        raise ScheduleFormatError(
            "В файле недостаточно столбцов для дня, пары, времени и недели."
        )

    data_start_idx = group_row_idx + 2
    if data_start_idx >= len(df.index):
        raise ScheduleFormatError(
            "После строки с группами не найдены строки расписания."
        )

    schedule_rows = df.iloc[data_start_idx:].copy()
    with pd.option_context("future.no_silent_downcasting", True):
        common_values = schedule_rows.iloc[:, :COMMON_COLUMNS_COUNT].ffill()
    schedule_rows.iloc[:, :COMMON_COLUMNS_COUNT] = common_values
    return schedule_rows


def _safe_cell(df: pd.DataFrame, row_idx: int, col_idx: int) -> str:
    """Читает ячейку по позиции, не падая на обрезанном хвосте таблицы."""

    if col_idx < 0 or col_idx >= df.shape[1]:
        return ""
    return clean_cell(df.iat[row_idx, col_idx])


def _pair_sort_key(value: Any) -> float:
    """Извлекает числовую часть пары; неизвестные значения отправляет в конец."""

    match = PAIR_NUMBER_PATTERN.search(clean_cell(value))
    if match is None:
        return math.inf

    try:
        return float(match.group(0).replace(",", "."))
    except ValueError:
        return math.inf


def _day_sort_key(value: Any) -> int:
    """Возвращает логический номер дня недели."""

    normalized_day = clean_cell(value).casefold().rstrip(".")
    return DAY_ORDER.get(normalized_day, len(DAY_ORDER))


def _is_schedule_row(df: pd.DataFrame, row_idx: int) -> bool:
    """Отделяет строки занятий от подписей и примечаний внизу Excel-листа."""

    day_order = _day_sort_key(df.iat[row_idx, 0])
    pair_order = _pair_sort_key(df.iat[row_idx, 1])
    week = clean_cell(df.iat[row_idx, 4]).casefold()

    return (
        day_order <= 5
        and math.isfinite(pair_order)
        and _is_time_value(df.iat[row_idx, 2])
        and _is_time_value(df.iat[row_idx, 3])
        and week in {"i", "ii"}
    )


def sort_schedule(result: pd.DataFrame) -> pd.DataFrame:
    """Сортирует результат по дню и номеру пары, сохраняя порядок совпадений."""

    if result.empty:
        return result.reindex(columns=OUTPUT_COLUMNS).reset_index(drop=True)

    sortable = result.copy()
    sortable["__day_order"] = sortable["День недели"].map(_day_sort_key)
    sortable["__pair_order"] = sortable["Пара"].map(_pair_sort_key)
    sortable["__source_order"] = range(len(sortable.index))

    sortable = sortable.sort_values(
        by=["__day_order", "__pair_order", "__source_order"],
        kind="mergesort",
        na_position="last",
    )
    return sortable.drop(
        columns=["__day_order", "__pair_order", "__source_order"]
    ).reset_index(drop=True)


def sort_internal_schedule(
    result: pd.DataFrame,
    teacher_queries: Sequence[str],
) -> pd.DataFrame:
    """Сортирует пакетный результат с сохранением порядка введённых фамилий."""

    if result.empty:
        return result.reindex(columns=INTERNAL_COLUMNS).reset_index(drop=True)

    query_order = {
        _normalize_search_text(query): position
        for position, query in enumerate(teacher_queries)
    }
    sortable = result.copy()
    sortable["__query_order"] = sortable[QUERY_COLUMN].map(
        lambda value: query_order.get(_normalize_search_text(value), len(query_order))
    )
    sortable["__day_order"] = sortable["День недели"].map(_day_sort_key)
    sortable["__pair_order"] = sortable["Пара"].map(_pair_sort_key)
    sortable["__week_order"] = sortable["Неделя"].map(
        lambda value: 0 if clean_cell(value).casefold() == "i" else 1
    )
    sortable["__source_order"] = range(len(sortable.index))
    sortable = sortable.sort_values(
        by=[
            "__query_order",
            "__day_order",
            "__pair_order",
            "__week_order",
            "__source_order",
        ],
        kind="mergesort",
        na_position="last",
    )
    return sortable.drop(
        columns=[
            "__query_order",
            "__day_order",
            "__pair_order",
            "__week_order",
            "__source_order",
        ]
    ).reset_index(drop=True)


def _join_unique(values: Sequence[Any], separator: str = ", ") -> str:
    """Объединяет непустые значения без дублей, сохраняя исходный порядок."""

    unique_values: list[str] = []
    for value in values:
        cleaned_value = clean_cell(value)
        if cleaned_value and cleaned_value not in unique_values:
            unique_values.append(cleaned_value)
    return separator.join(unique_values)


def _lesson_count_text(count: int) -> str:
    """Возвращает согласованную русскую подпись количества занятий."""

    last_two_digits = count % 100
    last_digit = count % 10
    if 11 <= last_two_digits <= 14:
        noun = "занятий"
    elif last_digit == 1:
        noun = "занятие"
    elif 2 <= last_digit <= 4:
        noun = "занятия"
    else:
        noun = "занятий"
    return f"{count} {noun}"


def _teacher_conflict_identity(query: Any, teacher: Any) -> str:
    """Определяет преподавателя для сравнения занятий одного временного слота."""

    normalized_query = _normalize_search_text(query)
    if len(normalized_query.split()) >= 2:
        # Полное имя или фамилия с инициалами надёжнее плавающей записи в Excel.
        return normalized_query
    return _normalize_search_text(teacher) or normalized_query


def collapse_schedule_conflicts(records: pd.DataFrame) -> pd.DataFrame:
    """Сворачивает занятия слота и отмечает только разные занятия преподавателя."""

    if records.empty:
        return pd.DataFrame(columns=DISPLAY_COLUMNS)

    required_columns = set(INTERNAL_COLUMNS)
    if missing_columns := required_columns.difference(records.columns):
        raise ValueError(
            "Для поиска накладок не хватает столбцов: "
            + ", ".join(sorted(missing_columns))
        )

    display_records: list[dict[str, str]] = []
    grouped_records = records.copy()
    identity_column = "__teacher_conflict_identity"
    grouped_records[identity_column] = [
        _teacher_conflict_identity(query, teacher)
        for query, teacher in zip(
            grouped_records[QUERY_COLUMN],
            grouped_records[TEACHER_COLUMN],
            strict=True,
        )
    ]
    slot_columns = [
        QUERY_COLUMN,
        identity_column,
        "День недели",
        "Пара",
        "Неделя",
    ]

    for _, slot_records in grouped_records.groupby(
        slot_columns,
        sort=False,
        dropna=False,
    ):
        logical_lessons: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        for record in slot_records.to_dict(orient="records"):
            lesson_values = (
                clean_cell(record["Время"]),
                clean_cell(record["Дисциплина"]),
                clean_cell(record["Вид занятий"]),
                clean_cell(record["Аудитория"]),
            )
            lesson_key = tuple(_normalize_search_text(value) for value in lesson_values)
            lesson = logical_lessons.setdefault(
                lesson_key,
                {
                    "time": lesson_values[0],
                    "discipline": lesson_values[1],
                    "lesson_type": lesson_values[2],
                    "auditorium": lesson_values[3],
                    "groups": [],
                    "sources": [],
                },
            )
            group = clean_cell(record["Группа"])
            source = clean_cell(record[SOURCE_COLUMN])
            if group and group not in lesson["groups"]:
                lesson["groups"].append(group)
            if source and source not in lesson["sources"]:
                lesson["sources"].append(source)

        lessons = list(logical_lessons.values())
        first_record = slot_records.iloc[0]
        conflict_text = (
            f"⚠ {_lesson_count_text(len(lessons))} одновременно"
            if len(lessons) > 1
            else ""
        )
        display_records.append(
            {
                QUERY_COLUMN: clean_cell(first_record[QUERY_COLUMN]),
                TEACHER_COLUMN: _join_unique(
                    slot_records[TEACHER_COLUMN].tolist(),
                    separator="\n",
                ),
                "День недели": clean_cell(first_record["День недели"]),
                "Пара": clean_cell(first_record["Пара"]),
                "Время": "\n".join(lesson["time"] for lesson in lessons),
                "Неделя": clean_cell(first_record["Неделя"]),
                CONFLICT_COLUMN: conflict_text,
                "Группа": "\n".join(
                    _join_unique(lesson["groups"]) for lesson in lessons
                ),
                "Дисциплина": "\n".join(lesson["discipline"] for lesson in lessons),
                "Вид занятий": "\n".join(lesson["lesson_type"] for lesson in lessons),
                "Аудитория": "\n".join(lesson["auditorium"] for lesson in lessons),
                SOURCE_COLUMN: "\n".join(
                    _join_unique(lesson["sources"]) for lesson in lessons
                ),
            }
        )

    display_frame = pd.DataFrame.from_records(
        display_records,
        columns=DISPLAY_COLUMNS,
    )
    return sort_internal_schedule(
        display_frame,
        parse_teacher_queries(records[QUERY_COLUMN].tolist()),
    )


def extract_schedule_records(
    df: pd.DataFrame,
    source_name: str = "",
) -> pd.DataFrame:
    """Извлекает все заполненные занятия, включая строки без преподавателя."""

    group_row_idx = find_group_row(df)
    schedule_rows = _prepare_schedule_rows(df, group_row_idx)

    if df.shape[1] <= FIRST_GROUP_COLUMN + 2:
        raise ScheduleFormatError("В файле не найдены столбцы групп и преподавателей.")

    records: list[dict[str, str]] = []
    group_columns = range(
        FIRST_GROUP_COLUMN,
        df.shape[1] - 2,
        GROUP_BLOCK_SIZE,
    )

    for row_idx in range(len(schedule_rows.index)):
        if not _is_schedule_row(schedule_rows, row_idx):
            continue

        day = _safe_cell(schedule_rows, row_idx, 0)
        pair = _format_pair(schedule_rows.iat[row_idx, 1])
        lesson_time = _build_time_range(
            schedule_rows.iat[row_idx, 2],
            schedule_rows.iat[row_idx, 3],
        )
        week = _safe_cell(schedule_rows, row_idx, 4)

        for col_idx in group_columns:
            group = _safe_cell(df, group_row_idx, col_idx)
            if not group or GROUP_PATTERN.search(group) is None:
                continue

            discipline = _safe_cell(schedule_rows, row_idx, col_idx)
            lesson_type = _safe_cell(schedule_rows, row_idx, col_idx + 1)
            teacher = _safe_cell(schedule_rows, row_idx, col_idx + 2)
            auditorium = _safe_cell(schedule_rows, row_idx, col_idx + 3)
            if not any((discipline, lesson_type, teacher, auditorium)):
                continue
            records.append(
                {
                    TEACHER_COLUMN: teacher,
                    "День недели": day,
                    "Пара": pair,
                    "Время": lesson_time,
                    "Неделя": week,
                    "Группа": group,
                    "Дисциплина": discipline,
                    "Вид занятий": lesson_type,
                    "Аудитория": auditorium,
                    SOURCE_COLUMN: clean_cell(source_name),
                }
            )

    result = pd.DataFrame.from_records(records, columns=ALL_SCHEDULE_COLUMNS)
    if result.empty:
        return result

    sortable = result.copy()
    sortable["__day_order"] = sortable["День недели"].map(_day_sort_key)
    sortable["__pair_order"] = sortable["Пара"].map(_pair_sort_key)
    sortable["__week_order"] = sortable["Неделя"].map(
        lambda value: 0 if _normalize_search_text(value) == "i" else 1
    )
    sortable["__source_order"] = range(len(sortable.index))
    sortable = sortable.sort_values(
        by=["__day_order", "__pair_order", "__week_order", "__source_order"],
        kind="mergesort",
        na_position="last",
    )
    return sortable.drop(
        columns=["__day_order", "__pair_order", "__week_order", "__source_order"]
    ).reset_index(drop=True)


def _filter_schedule_records(
    records: pd.DataFrame,
    teacher_queries: Sequence[str],
) -> pd.DataFrame:
    """Фильтрует полный набор занятий по введённым фамилиям или частям ФИО."""

    queries = parse_teacher_queries(teacher_queries)
    if not queries or records.empty:
        return pd.DataFrame(columns=INTERNAL_COLUMNS)

    normalized_queries = [(query, _normalize_search_text(query)) for query in queries]
    filtered_records: list[dict[str, str]] = []
    for record in records.to_dict(orient="records"):
        normalized_teacher = _normalize_search_text(record[TEACHER_COLUMN])
        for query, normalized_query in normalized_queries:
            if normalized_query not in normalized_teacher:
                continue
            filtered_records.append({QUERY_COLUMN: query, **record})

    result = pd.DataFrame.from_records(filtered_records, columns=INTERNAL_COLUMNS)
    return sort_internal_schedule(result, queries)


def parse_schedule_multi(
    df: pd.DataFrame,
    teacher_queries: Sequence[str],
    source_name: str = "",
) -> pd.DataFrame:
    """Извлекает занятия нескольких преподавателей за один проход по листу."""

    queries = parse_teacher_queries(teacher_queries)
    if not queries:
        raise ValueError("Введите хотя бы одну фамилию преподавателя.")
    return _filter_schedule_records(
        extract_schedule_records(df, source_name),
        queries,
    )


def parse_schedule(df: pd.DataFrame, teacher_query: str) -> pd.DataFrame:
    """Совместимый однопользовательский интерфейс старого парсера."""

    result = parse_schedule_multi(df, [teacher_query])
    return sort_schedule(result.reindex(columns=OUTPUT_COLUMNS))


def _deduplicate_batch_records(
    records: pd.DataFrame,
    teacher_queries: Sequence[str],
) -> pd.DataFrame:
    """Удаляет полные дубли занятий, сохраняя разные группы и запросы."""

    if records.empty:
        return records.reindex(columns=INTERNAL_COLUMNS).reset_index(drop=True)

    deduplication_columns = [
        QUERY_COLUMN,
        TEACHER_COLUMN,
        *OUTPUT_COLUMNS,
    ]
    deduplicated = records.drop_duplicates(
        subset=deduplication_columns,
        keep="first",
        ignore_index=True,
    )
    return sort_internal_schedule(deduplicated, teacher_queries)


def _deduplicate_all_schedule_records(records: pd.DataFrame) -> pd.DataFrame:
    """Удаляет полные дубли из набора всех занятий для сверки нагрузки."""

    if records.empty:
        return records.reindex(columns=ALL_SCHEDULE_COLUMNS).reset_index(drop=True)
    return records.drop_duplicates(
        subset=[TEACHER_COLUMN, *OUTPUT_COLUMNS],
        keep="first",
        ignore_index=True,
    ).reindex(columns=ALL_SCHEDULE_COLUMNS)


def parse_schedule_files(
    files: Sequence[tuple[str, bytes]],
    teacher_queries: Sequence[str],
    session_state: MutableMapping[str, Any],
) -> BatchParseResult:
    """Продолжает пакетную обработку, даже если отдельный файл повреждён."""

    queries = parse_teacher_queries(teacher_queries)
    processed_files: list[str] = []
    duplicate_files: list[str] = []
    errors: list[BatchFileError] = []
    dataframes: list[pd.DataFrame] = []
    all_dataframes: list[pd.DataFrame] = []
    seen_digests: set[str] = set()

    for position, (raw_filename, file_content) in enumerate(files, start=1):
        filename = clean_cell(raw_filename) or f"Файл {position}"
        file_digest = sha256(file_content).hexdigest()
        if file_digest in seen_digests:
            duplicate_files.append(filename)
            continue
        seen_digests.add(file_digest)

        try:
            source_df = read_excel_for_session(file_content, session_state)
            all_parsed = extract_schedule_records(source_df, filename)
            parsed = _filter_schedule_records(all_parsed, queries)
        except ScheduleError as exc:
            errors.append(BatchFileError(filename, str(exc)))
            continue
        except Exception:
            LOGGER.exception("Непредвиденная ошибка файла %s", filename)
            errors.append(
                BatchFileError(
                    filename,
                    "Непредвиденная ошибка обработки файла.",
                )
            )
            continue

        processed_files.append(filename)
        dataframes.append(parsed)
        all_dataframes.append(all_parsed)

    if dataframes:
        combined = pd.concat(dataframes, ignore_index=True)
        records = _deduplicate_batch_records(combined, queries)
    else:
        records = pd.DataFrame(columns=INTERNAL_COLUMNS)

    if all_dataframes:
        all_records = _deduplicate_all_schedule_records(
            pd.concat(all_dataframes, ignore_index=True)
        )
    else:
        all_records = pd.DataFrame(columns=ALL_SCHEDULE_COLUMNS)

    return BatchParseResult(
        records=records,
        all_records=all_records,
        processed_files=tuple(processed_files),
        duplicate_files=tuple(duplicate_files),
        errors=tuple(errors),
    )


def _find_workload_columns(df: pd.DataFrame) -> tuple[int, dict[str, int]]:
    """Находит строку заголовков и обязательные столбцы файла нагрузки."""

    required_headers = {
        "subject": "мероприятие реестра",
        "lesson_type": "вид потока",
        "plan_group": "план. поток",
        "group": "группа",
        "semester": "семестр",
        "teacher": "ппс",
    }
    for row_idx in range(min(15, len(df.index))):
        normalized_headers = [
            _normalize_search_text(value) for value in df.iloc[row_idx].tolist()
        ]
        columns: dict[str, int] = {}
        for key, expected_header in required_headers.items():
            for column_idx, header in enumerate(normalized_headers):
                if header == expected_header or (
                    key == "subject" and header.startswith(expected_header)
                ):
                    columns[key] = column_idx
                    break
        if len(columns) == len(required_headers):
            return row_idx, columns

    raise WorkloadFormatError(
        "В файле нагрузки не найдены столбцы «Мероприятие реестра», "
        "«Вид потока», «Группа», «Семестр» и «ППС»."
    )


def parse_workload_file(
    file: bytes | bytearray | BinaryIO,
    academic_term: str,
) -> pd.DataFrame:
    """Читает нагрузку и оставляет ЛК/ПР нужной чётности семестра."""

    df = read_excel_file(file)
    header_row_idx, columns = _find_workload_columns(df)
    normalized_term = _normalize_search_text(academic_term)
    if normalized_term not in {"осень", "весна"}:
        raise ValueError("Учебный семестр должен быть «Осень» или «Весна».")
    expected_parity = 1 if normalized_term == "осень" else 0

    records: list[dict[str, Any]] = []
    for row_idx in range(header_row_idx + 1, len(df.index)):
        raw_subject = df.iat[row_idx, columns["subject"]]
        discipline = _display_workload_discipline(raw_subject)
        lesson_type = _normalize_lesson_type(df.iat[row_idx, columns["lesson_type"]])
        semester = _parse_semester_number(df.iat[row_idx, columns["semester"]])
        if (
            not discipline
            or lesson_type not in WORKLOAD_CLASS_TYPES
            or semester is None
            or semester % 2 != expected_parity
        ):
            continue

        teacher = clean_cell(df.iat[row_idx, columns["teacher"]])
        groups = _extract_groups(
            df.iat[row_idx, columns["plan_group"]],
            df.iat[row_idx, columns["group"]],
        ) or [""]
        for group in groups:
            records.append(
                {
                    "Преподаватель нагрузки": teacher,
                    "Дисциплина": discipline,
                    "Вид занятий": lesson_type,
                    "Группа": group,
                    "Семестр": semester,
                    "Строка нагрузки": row_idx + 1,
                }
            )

    workload = pd.DataFrame.from_records(records, columns=WORKLOAD_COLUMNS)
    if workload.empty:
        term_label = "нечётных" if expected_parity else "чётных"
        raise WorkloadFormatError(
            f"В нагрузке не найдены ЛК/ПР для {term_label} семестров."
        )
    return workload.drop_duplicates(
        subset=[
            "Преподаватель нагрузки",
            "Дисциплина",
            "Вид занятий",
            "Группа",
            "Семестр",
        ],
        keep="first",
        ignore_index=True,
    )


def _schedule_place(records: pd.DataFrame) -> str:
    """Формирует краткое описание возможного места в расписании."""

    places: list[str] = []
    for record in records.to_dict(orient="records"):
        place = (
            f"{clean_cell(record['День недели'])}, "
            f"пара {clean_cell(record['Пара'])}, "
            f"неделя {clean_cell(record['Неделя'])}, "
            f"{clean_cell(record['Время'])}; "
            f"{clean_cell(record['Группа'])}; "
            f"{clean_cell(record[SOURCE_COLUMN])}"
        )
        if place not in places:
            places.append(place)
    return "\n".join(places)


def _suggested_teacher_conflicts(
    blank_candidates: pd.DataFrame,
    schedule: pd.DataFrame,
    suggested_teacher: str,
    suggested_discipline_key: str,
) -> str:
    """Ищет занятия, конфликтующие с предлагаемой подстановкой ФИО."""

    surname = _teacher_surname_key(suggested_teacher)
    if not surname:
        return ""
    teacher_schedule = schedule[
        schedule[TEACHER_COLUMN].map(_teacher_surname_key) == surname
    ]
    conflicts: list[str] = []
    for candidate in blank_candidates.to_dict(orient="records"):
        same_slot = teacher_schedule[
            (
                teacher_schedule["День недели"].map(_day_sort_key)
                == _day_sort_key(candidate["День недели"])
            )
            & (
                teacher_schedule["Пара"].map(_format_pair)
                == _format_pair(candidate["Пара"])
            )
            & (
                teacher_schedule["Неделя"].map(_normalize_search_text)
                == _normalize_search_text(candidate["Неделя"])
            )
            & (teacher_schedule["__discipline_key"] != suggested_discipline_key)
        ]
        for conflict in same_slot.to_dict(orient="records"):
            description = (
                f"{clean_cell(conflict['День недели'])}, "
                f"пара {clean_cell(conflict['Пара'])}, "
                f"неделя {clean_cell(conflict['Неделя'])}: "
                f"{clean_cell(conflict['Дисциплина'])} "
                f"({clean_cell(conflict['Группа'])})"
            )
            if description not in conflicts:
                conflicts.append(description)
    return "\n".join(conflicts)


def audit_workload(
    workload: pd.DataFrame,
    schedule_records: pd.DataFrame,
) -> WorkloadAuditResult:
    """Сверяет назначения нагрузки с дисциплинами и ФИО в расписании."""

    if workload.empty:
        raise WorkloadFormatError("После фильтрации файл нагрузки пуст.")

    schedule = schedule_records.reindex(columns=ALL_SCHEDULE_COLUMNS).copy()
    schedule["__discipline_key"] = schedule["Дисциплина"].map(_normalize_discipline)
    schedule["__lesson_type_key"] = schedule["Вид занятий"].map(_normalize_lesson_type)
    schedule["__group_key"] = schedule["Группа"].map(
        lambda value: clean_cell(value).upper()
    )

    prepared_workload = workload.copy()
    prepared_workload["__discipline_key"] = prepared_workload["Дисциплина"].map(
        _normalize_discipline
    )
    assignment_columns = [
        "__discipline_key",
        "Вид занятий",
        "Группа",
        "Семестр",
    ]

    error_records: list[dict[str, str]] = []
    suggestion_records: list[dict[str, str]] = []
    matched_assignments = 0
    checked_assignments = 0

    for _, assignment in prepared_workload.groupby(
        assignment_columns,
        sort=False,
        dropna=False,
    ):
        checked_assignments += 1
        first = assignment.iloc[0]
        discipline = clean_cell(first["Дисциплина"])
        discipline_key = clean_cell(first["__discipline_key"])
        lesson_type = _normalize_lesson_type(first["Вид занятий"])
        group = clean_cell(first["Группа"]).upper()
        semester = clean_cell(first["Семестр"])
        expected_teachers = [
            teacher
            for teacher in dict.fromkeys(
                clean_cell(value)
                for value in assignment["Преподаватель нагрузки"].tolist()
            )
            if teacher
        ]
        expected_teacher_text = ", ".join(expected_teachers)
        expected_surnames = {
            surname
            for surname in map(_teacher_surname_key, expected_teachers)
            if surname
        }

        base_issue = {
            "Преподаватели по нагрузке": expected_teacher_text,
            "Дисциплина": discipline,
            "Вид занятий": lesson_type,
            "Группа": group,
            "Семестр": semester,
            "ФИО в расписании": "",
            "Предлагаемое ФИО": "",
            "Возможная накладка": "",
            "Возможное место в расписании": "",
        }
        if not expected_teachers:
            error_records.append(
                {
                    "Статус": "❌ Ошибка нагрузки",
                    "Проблема": "В нагрузке не указан преподаватель",
                    **base_issue,
                }
            )
            continue
        if not group:
            error_records.append(
                {
                    "Статус": "❌ Ошибка нагрузки",
                    "Проблема": "В нагрузке не указана группа",
                    **base_issue,
                }
            )
            continue

        subject_candidates = schedule[schedule["__discipline_key"] == discipline_key]
        exact_candidates = subject_candidates[
            (subject_candidates["__lesson_type_key"] == lesson_type)
            & (subject_candidates["__group_key"] == group)
        ]
        if exact_candidates.empty:
            problem = (
                "Дисциплина отсутствует в расписании"
                if subject_candidates.empty
                else "Не найден нужный вид занятия для группы"
            )
            error_records.append(
                {
                    "Статус": "❌ Не найдено",
                    "Проблема": problem,
                    **base_issue,
                    "Возможное место в расписании": _schedule_place(
                        subject_candidates.head(5)
                    ),
                }
            )
            continue

        scheduled_teachers = [
            teacher
            for teacher in dict.fromkeys(
                clean_cell(value) for value in exact_candidates[TEACHER_COLUMN].tolist()
            )
            if teacher
        ]
        scheduled_surnames = {
            surname
            for surname in map(_teacher_surname_key, scheduled_teachers)
            if surname
        }
        if expected_surnames and expected_surnames.issubset(scheduled_surnames):
            matched_assignments += 1
            continue

        blank_candidates = exact_candidates[
            exact_candidates[TEACHER_COLUMN].map(clean_cell) == ""
        ]
        if not blank_candidates.empty:
            suggested_teacher = (
                expected_teachers[0] if len(expected_teachers) == 1 else ""
            )
            possible_conflict = _suggested_teacher_conflicts(
                blank_candidates,
                schedule,
                suggested_teacher,
                discipline_key,
            )
            suggestion_records.append(
                {
                    "Статус": (
                        "⚠️ Возможная накладка"
                        if possible_conflict
                        else "⚠️ Возможное место"
                    ),
                    "Проблема": (
                        "После подстановки ФИО найден другой предмет в то же время"
                        if possible_conflict
                        else "В расписании не указано ФИО"
                    ),
                    **base_issue,
                    "ФИО в расписании": ", ".join(scheduled_teachers),
                    "Предлагаемое ФИО": suggested_teacher,
                    "Возможная накладка": possible_conflict,
                    "Возможное место в расписании": _schedule_place(blank_candidates),
                }
            )
            continue

        error_records.append(
            {
                "Статус": "❌ Несовпадение ФИО",
                "Проблема": "В расписании указан другой преподаватель",
                **base_issue,
                "ФИО в расписании": ", ".join(scheduled_teachers),
                "Возможное место в расписании": _schedule_place(exact_candidates),
            }
        )

    errors = pd.DataFrame.from_records(error_records, columns=WORKLOAD_ISSUE_COLUMNS)
    suggestions = pd.DataFrame.from_records(
        suggestion_records,
        columns=WORKLOAD_ISSUE_COLUMNS,
    )
    return WorkloadAuditResult(
        checked_assignments=checked_assignments,
        matched_assignments=matched_assignments,
        errors=errors,
        suggestions=suggestions,
    )


def _pair_number(value: Any) -> int | None:
    """Возвращает положительный целочисленный номер пары для экспортной сетки."""

    pair_value = _pair_sort_key(value)
    if not math.isfinite(pair_value) or not pair_value.is_integer():
        return None
    pair_number = int(pair_value)
    return pair_number if pair_number > 0 else None


def _week_order(value: Any) -> int | None:
    """Преобразует I/II в индексы 0/1."""

    normalized_week = clean_cell(value).casefold()
    if normalized_week == "i":
        return 0
    if normalized_week == "ii":
        return 1
    return None


def _teacher_display_name(records: pd.DataFrame, query: str) -> str:
    """Подставляет полное ФИО, если запрос однозначно ему соответствует."""

    query_records = records[
        records[QUERY_COLUMN].map(_normalize_search_text)
        == _normalize_search_text(query)
    ]
    candidates: list[str] = []
    for teacher_value in query_records[TEACHER_COLUMN].tolist():
        teacher = clean_cell(teacher_value)
        parts = [
            clean_cell(part)
            for part in TEACHER_NAME_SEPARATOR_PATTERN.split(teacher)
            if clean_cell(part)
        ]
        matching_parts = [
            part
            for part in parts
            if _normalize_search_text(query) in _normalize_search_text(part)
        ]
        candidates.extend(matching_parts or ([teacher] if teacher else []))

    if not candidates:
        return clean_cell(query)

    surname_keys = {
        _normalize_search_text(candidate.split(maxsplit=1)[0])
        for candidate in candidates
        if candidate.split(maxsplit=1)
    }
    if len(surname_keys) > 1:
        return clean_cell(query)

    counts = Counter(candidates)
    return max(counts, key=lambda candidate: (counts[candidate], len(candidate)))


def _export_clusters_for_query(
    records: pd.DataFrame,
    query: str,
) -> list[dict[str, Any]]:
    """Собирает группы совместного занятия, не смешивая разные занятия слота."""

    query_records = records[
        records[QUERY_COLUMN].map(_normalize_search_text)
        == _normalize_search_text(query)
    ]
    clusters: dict[tuple[Any, ...], dict[str, Any]] = {}

    for record in query_records.to_dict(orient="records"):
        day_order = _day_sort_key(record["День недели"])
        pair_number = _pair_number(record["Пара"])
        week_order = _week_order(record["Неделя"])
        if day_order > 5 or pair_number is None or week_order is None:
            continue

        key = (
            day_order,
            pair_number,
            week_order,
            _teacher_conflict_identity(query, record[TEACHER_COLUMN]),
            _normalize_search_text(record["Время"]),
            _normalize_search_text(record["Дисциплина"]),
            _normalize_search_text(record["Вид занятий"]),
            _normalize_search_text(record["Аудитория"]),
        )
        cluster = clusters.setdefault(
            key,
            {
                "day_order": day_order,
                "pair_number": pair_number,
                "week_order": week_order,
                "teacher_identity": key[3],
                "teacher": clean_cell(record[TEACHER_COLUMN]),
                "time": clean_cell(record["Время"]),
                "discipline": clean_cell(record["Дисциплина"]),
                "lesson_type": clean_cell(record["Вид занятий"]),
                "auditorium": clean_cell(record["Аудитория"]),
                "groups": [],
            },
        )
        group = clean_cell(record["Группа"])
        if group and group not in cluster["groups"]:
            cluster["groups"].append(group)

    result = list(clusters.values())
    result.sort(
        key=lambda cluster: (
            cluster["day_order"],
            cluster["pair_number"],
            cluster["week_order"],
            cluster["discipline"].casefold(),
            ", ".join(cluster["groups"]).casefold(),
            cluster["auditorium"].casefold(),
        )
    )
    return result


def build_schedule_xlsx(
    records: pd.DataFrame,
    teacher_queries: Sequence[str],
) -> bytes:
    """Создаёт цветную матрицу расписания в формате XLSX по образцу."""

    queries = parse_teacher_queries(teacher_queries)
    if not queries:
        raise ScheduleExportError("Не выбраны преподаватели для выгрузки.")
    if records.empty:
        raise ScheduleExportError("Нет найденных занятий для выгрузки.")

    try:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Расписание"
        worksheet.sheet_view.showGridLines = False
        worksheet.sheet_view.zoomScale = 80
        worksheet.freeze_panes = "D3"
        worksheet.print_title_rows = "1:2"
        worksheet.print_title_cols = "A:C"
        worksheet.page_setup.orientation = "landscape"
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        workbook.properties.title = "Расписание преподавателей"
        workbook.properties.creator = "Streamlit Schedule Parser"

        clusters_by_query = {
            _normalize_search_text(query): _export_clusters_for_query(records, query)
            for query in queries
        }
        maximum_pair = max(
            7,
            max(
                (
                    cluster["pair_number"]
                    for clusters in clusters_by_query.values()
                    for cluster in clusters
                ),
                default=7,
            ),
        )
        total_columns = 3 + len(queries) * 4
        rows_per_day = maximum_pair * 2
        total_rows = 2 + len(DAY_EXPORT_CONFIG) * rows_per_day

        thin_gray = Side(style="thin", color="A6A6A6")
        medium_black = Side(style="medium", color="000000")
        medium_red = Side(style="medium", color="C00000")
        header_fill = PatternFill("solid", fgColor="D3DFE2")
        teacher_fill = PatternFill("solid", fgColor="75FB4C")
        white_fill = PatternFill("solid", fgColor="FFFFFF")

        teacher_starts = {
            4 + teacher_index * 4 for teacher_index in range(len(queries))
        }
        teacher_ends = {7 + teacher_index * 4 for teacher_index in range(len(queries))}
        medium_left_columns = {1, *teacher_starts}
        medium_right_columns = {3, *teacher_ends, total_columns}
        day_start_rows = {
            3 + day_index * rows_per_day for day_index in range(len(DAY_EXPORT_CONFIG))
        }
        day_end_rows = {
            2 + (day_index + 1) * rows_per_day
            for day_index in range(len(DAY_EXPORT_CONFIG))
        }

        def cell_border(
            row: int,
            column: int,
            *,
            conflict: bool = False,
        ) -> Border:
            if conflict:
                return Border(
                    left=medium_red,
                    right=medium_red,
                    top=medium_red,
                    bottom=medium_red,
                )
            return Border(
                left=medium_black if column in medium_left_columns else thin_gray,
                right=medium_black if column in medium_right_columns else thin_gray,
                top=(medium_black if row == 1 or row in day_start_rows else thin_gray),
                bottom=(
                    medium_black
                    if row == 2 or row in day_end_rows or row == total_rows
                    else thin_gray
                ),
            )

        for column, value in enumerate(("день", "№", "Нед"), start=1):
            worksheet.merge_cells(
                start_row=1,
                start_column=column,
                end_row=2,
                end_column=column,
            )
            worksheet.cell(1, column, value)
            for row in (1, 2):
                cell = worksheet.cell(row, column)
                cell.fill = header_fill
                cell.font = Font(name="Arial", size=12, bold=True)
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

        for teacher_index, query in enumerate(queries):
            start_column = 4 + teacher_index * 4
            end_column = start_column + 3
            clusters = clusters_by_query[_normalize_search_text(query)]
            display_name = _teacher_display_name(records, query)
            header_value = f"{display_name}\n(занятий: {len(clusters)})"
            worksheet.merge_cells(
                start_row=1,
                start_column=start_column,
                end_row=1,
                end_column=end_column,
            )
            worksheet.cell(1, start_column, header_value)

            for column in range(start_column, end_column + 1):
                header_cell = worksheet.cell(1, column)
                header_cell.fill = teacher_fill
                header_cell.font = Font(name="Arial", size=14, bold=True)
                header_cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

            for offset, subheader in enumerate(EXPORT_SUBHEADERS):
                subheader_cell = worksheet.cell(2, start_column + offset, subheader)
                subheader_cell.fill = white_fill
                subheader_cell.font = Font(name="Arial", size=11, bold=False)
                subheader_cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

        slot_clusters: dict[tuple[str, int, int, int], list[dict[str, Any]]] = {}
        for query in queries:
            normalized_query = _normalize_search_text(query)
            for cluster in clusters_by_query[normalized_query]:
                slot_key = (
                    normalized_query,
                    cluster["day_order"],
                    cluster["pair_number"],
                    cluster["week_order"],
                )
                slot_clusters.setdefault(slot_key, []).append(cluster)

        for day_index, (day_label, day_color, odd_color, even_color) in enumerate(
            DAY_EXPORT_CONFIG
        ):
            day_start_row = 3 + day_index * rows_per_day
            day_end_row = day_start_row + rows_per_day - 1
            worksheet.cell(day_start_row, 1, day_label)

            for pair_number in range(1, maximum_pair + 1):
                pair_start_row = day_start_row + (pair_number - 1) * 2
                worksheet.cell(pair_start_row, 2, pair_number)

                for week_order in (0, 1):
                    row = pair_start_row + week_order
                    worksheet.cell(row, 3, week_order + 1)
                    body_color = odd_color if pair_number % 2 else even_color
                    body_fill = PatternFill("solid", fgColor=body_color)
                    day_fill = PatternFill("solid", fgColor=day_color)
                    maximum_lines = 0

                    for column in range(1, total_columns + 1):
                        cell = worksheet.cell(row, column)
                        cell.fill = day_fill if column == 1 else body_fill
                        cell.font = Font(name="Arial", size=10)
                        cell.alignment = Alignment(
                            horizontal="center" if column <= 3 else "left",
                            vertical="center",
                            wrap_text=True,
                        )

                    for teacher_index, query in enumerate(queries):
                        start_column = 4 + teacher_index * 4
                        clusters = slot_clusters.get(
                            (
                                _normalize_search_text(query),
                                day_index,
                                pair_number,
                                week_order,
                            ),
                            [],
                        )
                        lesson_counts_by_teacher = Counter(
                            cluster["teacher_identity"] for cluster in clusters
                        )
                        conflicting_teachers = {
                            teacher_identity
                            for teacher_identity, lesson_count in (
                                lesson_counts_by_teacher.items()
                            )
                            if lesson_count > 1
                        }
                        has_overlap = bool(conflicting_teachers)
                        maximum_lines = max(maximum_lines, len(clusters))
                        disciplines = [cluster["discipline"] for cluster in clusters]
                        if has_overlap and disciplines:
                            warning_index = next(
                                index
                                for index, cluster in enumerate(clusters)
                                if cluster["teacher_identity"] in conflicting_teachers
                            )
                            disciplines[warning_index] = (
                                f"⚠ {disciplines[warning_index]}"
                            )
                        values = (
                            "\n".join(disciplines),
                            "\n".join(cluster["lesson_type"] for cluster in clusters),
                            "\n".join(
                                ", ".join(cluster["groups"]) for cluster in clusters
                            ),
                            "\n".join(cluster["auditorium"] for cluster in clusters),
                        )
                        for offset, value in enumerate(values):
                            data_cell = worksheet.cell(
                                row, start_column + offset, value
                            )
                            data_cell.alignment = Alignment(
                                horizontal="left" if offset == 0 else "center",
                                vertical="center",
                                wrap_text=True,
                            )
                            data_cell.border = cell_border(
                                row,
                                start_column + offset,
                                conflict=has_overlap,
                            )
                            if has_overlap and offset == 0:
                                data_cell.font = Font(
                                    name="Arial",
                                    size=10,
                                    bold=True,
                                    color="C00000",
                                )

                    if maximum_lines:
                        worksheet.row_dimensions[row].height = min(
                            100,
                            38 + (maximum_lines - 1) * 22,
                        )
                    else:
                        worksheet.row_dimensions[row].height = 24

            for row in range(day_start_row, day_end_row + 1):
                worksheet.cell(row, 1).fill = PatternFill("solid", fgColor=day_color)
            worksheet.merge_cells(
                start_row=day_start_row,
                start_column=1,
                end_row=day_end_row,
                end_column=1,
            )
            worksheet.cell(day_start_row, 1).font = Font(
                name="Arial",
                size=12,
                bold=True,
            )
            worksheet.cell(day_start_row, 1).alignment = Alignment(
                horizontal="center",
                vertical="center",
            )

            for pair_number in range(1, maximum_pair + 1):
                pair_start_row = day_start_row + (pair_number - 1) * 2
                worksheet.merge_cells(
                    start_row=pair_start_row,
                    start_column=2,
                    end_row=pair_start_row + 1,
                    end_column=2,
                )
                worksheet.cell(pair_start_row, 2).alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                )

        for row in range(1, total_rows + 1):
            for column in range(1, total_columns + 1):
                cell = worksheet.cell(row, column)
                if cell.border == Border():
                    cell.border = cell_border(row, column)

        worksheet.row_dimensions[1].height = 72
        worksheet.row_dimensions[2].height = 26
        worksheet.column_dimensions["A"].width = 8
        worksheet.column_dimensions["B"].width = 5
        worksheet.column_dimensions["C"].width = 6
        teacher_widths = (32, 9, 16, 20)
        for teacher_index in range(len(queries)):
            start_column = 4 + teacher_index * 4
            for offset, width in enumerate(teacher_widths):
                worksheet.column_dimensions[
                    get_column_letter(start_column + offset)
                ].width = width

        worksheet.print_area = f"A1:{get_column_letter(total_columns)}{total_rows}"

        output = BytesIO()
        workbook.save(output)
        return output.getvalue()
    except ScheduleExportError:
        raise
    except Exception as exc:
        LOGGER.exception("Не удалось сформировать XLSX")
        raise ScheduleExportError(
            "Не удалось сформировать Excel-файл с расписанием."
        ) from exc


def _render_workload_audit(
    st: Any,
    audit: WorkloadAuditResult,
    academic_term: str,
) -> None:
    """Показывает результаты нагрузки отдельно от основного расписания."""

    st.subheader("Сверка с учебной нагрузкой")
    parity_text = (
        "нечётные" if _normalize_search_text(academic_term) == "осень" else "чётные"
    )
    st.caption(
        f"Проверены {parity_text} семестры. Предположения этого блока "
        "не добавляются в XLSX-выгрузку."
    )
    summary_columns = st.columns(4)
    summary_columns[0].metric("Назначений", audit.checked_assignments)
    summary_columns[1].metric("Совпало", audit.matched_assignments)
    summary_columns[2].metric("Ошибок", len(audit.errors.index))
    summary_columns[3].metric("Возможных мест", len(audit.suggestions.index))

    if audit.errors.empty and audit.suggestions.empty:
        st.success("Нагрузка соответствует загруженному расписанию.")
        return
    if not audit.errors.empty:
        st.error(
            f"Ошибки сверки нагрузки: {len(audit.errors.index)}. "
            "Эти строки требуют проверки исходных файлов."
        )
        st.dataframe(
            audit.errors,
            use_container_width=True,
            hide_index=True,
            row_height=54,
        )
    if not audit.suggestions.empty:
        possible_conflicts = int(
            audit.suggestions["Возможная накладка"].astype(bool).sum()
        )
        message = (
            f"Возможные места для отсутствующих ФИО: {len(audit.suggestions.index)}."
        )
        if possible_conflicts:
            message += f" Возможных накладок после подстановки: {possible_conflicts}."
        st.warning(message)
        st.dataframe(
            audit.suggestions,
            use_container_width=True,
            hide_index=True,
            row_height=70,
        )


def _render_footer(st: Any) -> None:
    """Показывает постоянный нижний блок со справкой и авторством."""

    st.divider()
    help_column, version_column, author_column = st.columns([1.2, 1.6, 7.2])
    with help_column:
        with st.popover("❓ Справка"):
            st.markdown("### Как пользоваться")
            st.markdown(
                """
1. Нажмите **Browse files** и выберите один или несколько файлов `.xls`/`.xlsx`.
2. При необходимости загрузите отдельный файл нагрузки и выберите **Осень** или **Весна**.
3. Введите фамилии преподавателей — каждую с новой строки либо через запятую.
4. Нажмите **Найти расписание**.
5. Проверьте расписание и отдельный блок сверки нагрузки. Жёлтый значок **⚠** показывает возможное место для отсутствующего ФИО; такие предположения не попадают в XLSX.
6. Нажмите **Скачать сводное расписание XLSX**, чтобы получить цветную Excel-таблицу.

**Как это работает:** приложение написано на Python и Streamlit. Pandas читает загруженные Excel-файлы, очищает объединённые и «грязные» ячейки, после чего поиск собирает занятия выбранных преподавателей из всех курсов.
                """
            )
    with version_column:
        with st.popover(f"Версия {APP_VERSION}"):
            st.markdown("### История изменений")
            for version, release_date, changes in APP_CHANGELOG:
                st.markdown(f"**{version} — {release_date}**")
                for change in changes:
                    st.markdown(f"- {change}")
    with author_column:
        st.caption("Автор: Трушин С.")


def _render_app(st: Any) -> None:
    """Отрисовывает основную часть пользовательского интерфейса."""

    st.title("Расписание преподавателя")
    st.caption(
        "Загрузите расписания нескольких курсов, укажите нужных преподавателей "
        "и скачайте сводную цветную сетку."
    )

    uploaded_files = st.file_uploader(
        "Excel-файлы расписания",
        type=["xls", "xlsx"],
        accept_multiple_files=True,
        help=(
            "Можно выбрать сразу несколько файлов .xls/.xlsx; "
            "в каждом обрабатывается первый лист."
        ),
        key="schedule_files",
    )
    academic_term = st.radio(
        "Текущий учебный семестр",
        options=("Осень", "Весна"),
        horizontal=True,
        help=("Для осени проверяются нечётные семестры нагрузки, для весны — чётные."),
    )
    workload_file = st.file_uploader(
        "Файл учебной нагрузки — необязательно",
        type=["xls", "xlsx"],
        accept_multiple_files=False,
        help=("Если файл не загружен, приложение работает только с расписанием."),
        key="workload_file",
    )
    teacher_input = st.text_area(
        "Преподаватели",
        placeholder="Иванов\nПетров\nСидорова",
        help=(
            "Введите фамилии или части ФИО с новой строки, через запятую "
            "или точку с запятой."
        ),
        height=120,
    )
    button_clicked = st.button("Найти расписание", type="primary")
    search_requested = button_clicked or bool(teacher_input.strip())

    if not search_requested:
        return
    if not uploaded_files:
        st.info("Сначала загрузите хотя бы один Excel-файл с расписанием.")
        return
    teacher_queries = parse_teacher_queries(teacher_input)
    if not teacher_queries and workload_file is None:
        st.warning(
            "Введите хотя бы одну фамилию преподавателя или загрузите файл нагрузки."
        )
        return

    with st.spinner("Читаю файл и ищу занятия…"):
        try:
            batch_result = parse_schedule_files(
                [
                    (uploaded_file.name, uploaded_file.getvalue())
                    for uploaded_file in uploaded_files
                ],
                teacher_queries,
                st.session_state,
            )
        except (ScheduleError, ValueError) as exc:
            st.error(str(exc))
            return
        except Exception:
            LOGGER.exception("Непредвиденная ошибка обработки расписания")
            st.error(
                "Не удалось обработать расписание из-за непредвиденной ошибки. "
                "Проверьте структуру файла и попробуйте снова."
            )
            return

    workload_audit: WorkloadAuditResult | None = None
    workload_error = ""
    if workload_file is not None:
        with st.spinner("Сверяю учебную нагрузку с расписанием…"):
            try:
                workload = parse_workload_file(
                    workload_file.getvalue(),
                    academic_term,
                )
                workload_audit = audit_workload(
                    workload,
                    batch_result.all_records,
                )
            except (ScheduleError, ValueError) as exc:
                workload_error = str(exc)
            except Exception:
                LOGGER.exception("Непредвиденная ошибка сверки нагрузки")
                workload_error = (
                    "Не удалось сверить нагрузку из-за непредвиденной ошибки."
                )

    if batch_result.errors:
        st.warning("Некоторые файлы не обработаны. Остальные результаты сохранены.")
        with st.expander("Ошибки файлов"):
            for file_error in batch_result.errors:
                st.write(f"• {file_error.filename}: {file_error.message}")

    if batch_result.duplicate_files:
        duplicate_names = ", ".join(batch_result.duplicate_files)
        st.info(f"Повторно загруженные файлы пропущены: {duplicate_names}")

    if not batch_result.processed_files:
        st.error("Не удалось обработать ни одного загруженного файла.")
        return

    if workload_error:
        st.error(f"Ошибка файла нагрузки: {workload_error}")
    elif workload_audit is not None:
        _render_workload_audit(st, workload_audit, academic_term)

    result_df = batch_result.records
    if not teacher_queries:
        st.info(
            "Сверка нагрузки завершена. Для вывода персонального расписания "
            "введите фамилию преподавателя."
        )
        return
    if result_df.empty:
        st.warning("Занятия указанных преподавателей не найдены.")
        return

    found_queries = {
        _normalize_search_text(value) for value in result_df[QUERY_COLUMN].tolist()
    }
    missing_queries = [
        query
        for query in teacher_queries
        if _normalize_search_text(query) not in found_queries
    ]
    st.success(
        f"Найдено записей: {len(result_df)}; "
        f"обработано файлов: {len(batch_result.processed_files)}."
    )
    if missing_queries:
        st.info("Без совпадений: " + ", ".join(missing_queries))

    display_df = collapse_schedule_conflicts(result_df)
    conflict_count = int(display_df[CONFLICT_COLUMN].astype(bool).sum())
    if conflict_count:
        st.warning(
            f"Обнаружено накладок: {conflict_count}. Разные занятия одного "
            "временного слота показаны вместе в одной строке."
        )

    st.dataframe(
        display_df.reindex(columns=DISPLAY_COLUMNS),
        use_container_width=True,
        hide_index=True,
        row_height=54 if conflict_count else None,
    )

    try:
        export_bytes = build_schedule_xlsx(result_df, teacher_queries)
    except ScheduleError as exc:
        st.error(str(exc))
        return

    st.download_button(
        "Скачать сводное расписание XLSX",
        data=export_bytes,
        file_name="расписание_преподавателей.xlsx",
        mime=("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        type="primary",
    )


def main() -> None:
    """Запускает постоянно доступный пользовательский интерфейс Streamlit."""

    try:
        import streamlit as st
    except ModuleNotFoundError as exc:  # pragma: no cover - подсказка для CLI-запуска
        raise RuntimeError(
            "Streamlit не установлен. Выполните 'pip install -r requirements.txt', "
            "затем запустите 'streamlit run app.py'."
        ) from exc

    st.set_page_config(
        page_title="Поиск расписания преподавателя",
        layout="wide",
    )
    try:
        _render_app(st)
    finally:
        _render_footer(st)


if __name__ == "__main__":
    main()
