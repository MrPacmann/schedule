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

HEADER_SEARCH_ROWS: Final[int] = 10
COMMON_COLUMNS_COUNT: Final[int] = 5
FIRST_GROUP_COLUMN: Final[int] = 5
GROUP_BLOCK_SIZE: Final[int] = 5
SESSION_CACHE_KEY: Final[str] = "_schedule_excel_cache"
MAX_CACHED_FILES: Final[int] = 20

QUERY_COLUMN: Final[str] = "Искомый преподаватель"
TEACHER_COLUMN: Final[str] = "Преподаватель"
SOURCE_COLUMN: Final[str] = "Файл-источник"
CONFLICT_COLUMN: Final[str] = "Накладка"

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


@dataclass(frozen=True, slots=True)
class BatchFileError:
    """Ошибка одного файла внутри пакетной обработки."""

    filename: str
    message: str


@dataclass(frozen=True, slots=True)
class BatchParseResult:
    """Результат отказоустойчивой обработки набора Excel-файлов."""

    records: pd.DataFrame
    processed_files: tuple[str, ...]
    duplicate_files: tuple[str, ...]
    errors: tuple[BatchFileError, ...]


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


def _group_count_text(count: int) -> str:
    """Возвращает согласованную русскую подпись количества групп."""

    last_two_digits = count % 100
    last_digit = count % 10
    if 11 <= last_two_digits <= 14:
        noun = "групп"
    elif last_digit == 1:
        noun = "группа"
    elif 2 <= last_digit <= 4:
        noun = "группы"
    else:
        noun = "групп"
    return f"{count} {noun}"


def collapse_schedule_conflicts(records: pd.DataFrame) -> pd.DataFrame:
    """Сворачивает повторные назначения одного слота в одну строку."""

    if records.empty:
        return pd.DataFrame(columns=DISPLAY_COLUMNS)

    required_columns = set(INTERNAL_COLUMNS)
    if missing_columns := required_columns.difference(records.columns):
        raise ValueError(
            "Для поиска накладок не хватает столбцов: "
            + ", ".join(sorted(missing_columns))
        )

    display_records: list[dict[str, str]] = []
    slot_columns = [
        QUERY_COLUMN,
        TEACHER_COLUMN,
        "День недели",
        "Пара",
        "Неделя",
    ]

    for _, slot_records in records.groupby(
        slot_columns,
        sort=False,
        dropna=False,
    ):
        logical_lessons: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        for record in slot_records.to_dict(orient="records"):
            lesson_key = (
                clean_cell(record["Время"]),
                clean_cell(record["Дисциплина"]),
                clean_cell(record["Вид занятий"]),
                clean_cell(record["Аудитория"]),
            )
            lesson = logical_lessons.setdefault(
                lesson_key,
                {
                    "time": lesson_key[0],
                    "discipline": lesson_key[1],
                    "lesson_type": lesson_key[2],
                    "auditorium": lesson_key[3],
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

        assignment_count = len(slot_records.index)
        if assignment_count <= 1:
            for record in slot_records.to_dict(orient="records"):
                display_record = {
                    column: clean_cell(record.get(column, ""))
                    for column in INTERNAL_COLUMNS
                }
                display_record[CONFLICT_COLUMN] = ""
                display_records.append(display_record)
            continue

        lessons = list(logical_lessons.values())
        first_record = slot_records.iloc[0]
        if len(lessons) > 1:
            conflict_text = f"⚠ {_lesson_count_text(len(lessons))} одновременно"
        else:
            distinct_group_count = len(lessons[0]["groups"])
            conflict_text = (
                "⚠ "
                f"{_group_count_text(distinct_group_count or assignment_count)} "
                "одновременно"
            )
        display_records.append(
            {
                QUERY_COLUMN: clean_cell(first_record[QUERY_COLUMN]),
                TEACHER_COLUMN: clean_cell(first_record[TEACHER_COLUMN]),
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


def parse_schedule_multi(
    df: pd.DataFrame,
    teacher_queries: Sequence[str],
    source_name: str = "",
) -> pd.DataFrame:
    """Извлекает занятия нескольких преподавателей за один проход по листу."""

    queries = parse_teacher_queries(teacher_queries)
    if not queries:
        raise ValueError("Введите хотя бы одну фамилию преподавателя.")

    normalized_queries = [(query, _normalize_search_text(query)) for query in queries]

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

            normalized_teacher = _normalize_search_text(teacher)
            matching_queries = [
                query
                for query, normalized_query in normalized_queries
                if normalized_query in normalized_teacher
            ]
            if not matching_queries:
                continue

            for query in matching_queries:
                records.append(
                    {
                        QUERY_COLUMN: query,
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

    result = pd.DataFrame.from_records(records, columns=INTERNAL_COLUMNS)
    return sort_internal_schedule(result, queries)


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


def parse_schedule_files(
    files: Sequence[tuple[str, bytes]],
    teacher_queries: Sequence[str],
    session_state: MutableMapping[str, Any],
) -> BatchParseResult:
    """Продолжает пакетную обработку, даже если отдельный файл повреждён."""

    queries = parse_teacher_queries(teacher_queries)
    if not queries:
        raise ValueError("Введите хотя бы одну фамилию преподавателя.")

    processed_files: list[str] = []
    duplicate_files: list[str] = []
    errors: list[BatchFileError] = []
    dataframes: list[pd.DataFrame] = []
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
            parsed = parse_schedule_multi(source_df, queries, filename)
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

    if dataframes:
        combined = pd.concat(dataframes, ignore_index=True)
        records = _deduplicate_batch_records(combined, queries)
    else:
        records = pd.DataFrame(columns=INTERNAL_COLUMNS)

    return BatchParseResult(
        records=records,
        processed_files=tuple(processed_files),
        duplicate_files=tuple(duplicate_files),
        errors=tuple(errors),
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
            clean_cell(record[TEACHER_COLUMN]),
            clean_cell(record["Время"]),
            clean_cell(record["Дисциплина"]),
            clean_cell(record["Вид занятий"]),
            clean_cell(record["Аудитория"]),
        )
        cluster = clusters.setdefault(
            key,
            {
                "day_order": day_order,
                "pair_number": pair_number,
                "week_order": week_order,
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
                        assignment_count = sum(
                            max(1, len(cluster["groups"])) for cluster in clusters
                        )
                        has_overlap = assignment_count > 1
                        maximum_lines = max(maximum_lines, len(clusters))
                        disciplines = [cluster["discipline"] for cluster in clusters]
                        if has_overlap and disciplines:
                            disciplines[0] = f"⚠ {disciplines[0]}"
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


def _render_footer(st: Any) -> None:
    """Показывает постоянный нижний блок со справкой и авторством."""

    st.divider()
    help_column, author_column = st.columns([1, 8])
    with help_column:
        with st.popover("❓ Справка"):
            st.markdown("### Как пользоваться")
            st.markdown(
                """
1. Нажмите **Browse files** и выберите один или несколько файлов `.xls`/`.xlsx`.
2. Введите фамилии преподавателей — каждую с новой строки либо через запятую.
3. Нажмите **Найти расписание**.
4. Проверьте найденные занятия в таблице. Значок **⚠** означает накладку: преподаватель одновременно назначен на несколько групп или занятий.
5. Нажмите **Скачать сводное расписание XLSX**, чтобы получить цветную Excel-таблицу.

**Как это работает:** приложение написано на Python и Streamlit. Pandas читает загруженные Excel-файлы, очищает объединённые и «грязные» ячейки, после чего поиск собирает занятия выбранных преподавателей из всех курсов.
                """
            )
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
    if not teacher_queries:
        st.warning("Введите хотя бы одну фамилию преподавателя.")
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

    result_df = batch_result.records
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
            f"Обнаружено накладок: {conflict_count}. Повторные назначения "
            "одного временного слота показаны вместе в одной строке."
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
