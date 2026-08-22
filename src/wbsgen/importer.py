"""記入済み Excel の読み込み。

このツールが作った空の WBS に書き込んだものを読み戻す。
見出しの文字で列を探すので、列を足したり順を入れ替えたりしていても読める。
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

from openpyxl import load_workbook

from .blank import MAX_ROWS, BlankWBS, SpecError
from .i18n import DEFAULT_LANGUAGE, LABELS, message, normalize
from .timeline import UNIT_DAY, UNIT_WEEK, VALID_UNITS
from .workcal import WEEKDAY_KEYS

def _build_header_maps():
    """全言語の見出しから、``見出し -> 属性名`` の表を作る。

    「開始」「日数」は予定と実績の両方にあるので、単独では決められない。
    それらは :data:`AMBIGUOUS_HEADERS` に回す。
    """
    plain: Dict[str, str] = {}
    ambiguous: Dict[str, tuple] = {}
    grouped: Dict[tuple, str] = {}

    for text in LABELS.values():
        pairs = [
            (text.group_plan, "start", text.columns["start"]),
            (text.group_plan, "days", text.columns["days"]),
            (text.group_plan, "end", text.columns["end"]),
            (text.group_actual, "actual_start", text.columns["actual_start"]),
            (text.group_actual, "actual_days", text.columns["actual_days"]),
            (text.group_actual, "actual_end", text.columns["actual_end"]),
        ]
        for group, key, label in pairs:
            grouped[(group, label)] = key

        for plan_key, actual_key in (("start", "actual_start"),
                                     ("days", "actual_days"),
                                     ("end", "actual_end")):
            for label in (text.columns[plan_key], text.columns[actual_key]):
                ambiguous[label] = (plan_key, actual_key)

        for key, label in text.columns.items():
            if label not in ambiguous:
                plain.setdefault(label, key)

    return plain, ambiguous, grouped


#: 見出しの文字 -> 属性名 (予定/実績で重ならない列)
HEADERS, AMBIGUOUS_HEADERS, GROUPED_HEADERS = _build_header_maps()

#: 「項目」の列を探すときに見る見出し (言語ごと)
NAME_LABELS = {text.columns["name"] for text in LABELS.values()}

#: 設定シートの表示単位の書き方 -> 内部の値
UNIT_LABELS = {
    label: unit
    for text in LABELS.values()
    for unit, label in text.units.items()
}

#: 設定シートの見出し -> 意味 (言語ごとの書き方をまとめる)
CONFIG_KEYS = {
    "start": {text.config_start for text in LABELS.values()},
    "period": {text.config_period for text in LABELS.values()},
    "unit": {text.config_unit for text in LABELS.values()},
}

#: 稼働日の行 (曜日名 -> weekday 番号) と「出」の書き方
WEEKDAY_LABELS = {
    name: index
    for text in LABELS.values()
    for index, name in enumerate(text.weekday_names)
}
WORK_ON_LABELS = {text.work_on for text in LABELS.values()}

#: シート名の候補 (言語ごと)
PLAN_SHEETS = {text.sheet_plan for text in LABELS.values()}
CONFIG_SHEETS = {text.sheet_config for text in LABELS.values()}
MEMBER_SHEETS = {text.sheet_member for text in LABELS.values()}
MEMBER_SKIP = ({text.member_title for text in LABELS.values()}
               | {text.member_headers[0] for text in LABELS.values()})

#: 見出しを探す範囲
HEADER_SEARCH_ROWS = 20
HEADER_SEARCH_COLS = 40


@dataclass
class Row:
    """記入された 1 行。読めた値をそのまま持つ。"""

    row: int = 0
    group: str = ""
    subgroup: str = ""
    no: str = ""
    name: str = ""
    start: Optional[_dt.date] = None
    days: Optional[int] = None
    end: Optional[_dt.date] = None
    actual_start: Optional[_dt.date] = None
    actual_days: Optional[int] = None
    actual_end: Optional[_dt.date] = None
    delay: Optional[int] = None
    progress: Optional[float] = None
    effort: Optional[float] = None
    predecessor: str = ""
    member: str = ""
    status: str = ""

    @property
    def has_plan(self) -> bool:
        return self.start is not None and self.end is not None

    @property
    def has_actual(self) -> bool:
        return self.actual_start is not None


@dataclass
class ImportedWBS:
    """読み込んだ WBS 一式。"""

    title: str
    spec: BlankWBS
    rows: List[Row] = field(default_factory=list)
    #: 読めなかった行の説明 (画面に出して知らせる)
    warnings: List[str] = field(default_factory=list)


# ----------------------------------------------------------------------
def read(source, filename: str = "", language: str = DEFAULT_LANGUAGE) -> ImportedWBS:
    """``.xlsx`` を読み込む。``source`` はパスかファイルオブジェクト。

    日本語版・英語版のどちらのファイルも読める。``language`` は
    メッセージと、読み込んだあとの表示に使う言語。
    """
    lang = normalize(language)
    try:
        book = load_workbook(source, data_only=True, read_only=False)
    except Exception as exc:  # noqa: BLE001 - openpyxl の例外は多岐にわたる
        raise SpecError(message(lang, "not_excel", reason=exc))

    sheet = _pick_sheet(book, PLAN_SHEETS)
    header_row, columns = _find_headers(sheet, lang)
    if "name" not in columns:
        raise SpecError(message(lang, "no_task_column"))

    config = _read_config(book)
    first_row = header_row + (2 if config.get("unit") == UNIT_DAY else 1)
    rows, warnings = _read_rows(sheet, first_row, columns, lang)

    spec = _build_spec(config, rows, book, filename, lang)
    title = _read_title(sheet, header_row) or spec.title
    spec.title = title
    return ImportedWBS(title=title, spec=spec, rows=rows, warnings=warnings)


def _pick_sheet(book, names):
    """言語ごとのシート名を順に探し、無ければ先頭のシートを使う。"""
    for name in book.sheetnames:
        if name in names:
            return book[name]
    return book.worksheets[0]


# ----------------------------------------------------------------------
def _find_headers(sheet, language: str) -> "tuple[int, Dict[str, int]]":
    """見出し行を探し、``属性名 -> 列番号`` を返す。"""
    max_row = min(sheet.max_row or 1, HEADER_SEARCH_ROWS)
    max_col = min(sheet.max_column or 1, HEADER_SEARCH_COLS)
    for row in range(1, max_row + 1):
        labels = {}
        for col in range(1, max_col + 1):
            value = sheet.cell(row=row, column=col).value
            if isinstance(value, str) and value.strip():
                labels[col] = value.strip()
        if not NAME_LABELS & set(labels.values()):
            continue
        return row, _map_columns(sheet, row, labels)
    raise SpecError(message(language, "no_header"))


def _map_columns(sheet, header_row: int, labels: Dict[int, str]) -> Dict[str, int]:
    """見出しを属性名に対応づける。

    「日数」「開始」などは予定と実績の両方にあるので、1 行上の帯
    (予定 / 実績) と組にして見分ける。帯が無い・ずれている場合は、
    左から出てきた順に「予定 → 実績」とみなす。
    """
    groups = _band_groups(sheet, header_row - 1, max(labels) if labels else 1)
    columns: Dict[str, int] = {}
    seen: Dict[str, int] = {}

    for col, label in sorted(labels.items()):
        key = GROUPED_HEADERS.get((groups.get(col, ""), label))
        if key is None and label in AMBIGUOUS_HEADERS:
            pair = AMBIGUOUS_HEADERS[label]
            index = min(seen.get(pair[0], 0), len(pair) - 1)
            key = pair[index]
        if key is None:
            key = HEADERS.get(label)
        if key is None or key in columns:
            continue
        columns[key] = col
        if label in AMBIGUOUS_HEADERS:
            first = AMBIGUOUS_HEADERS[label][0]
            seen[first] = seen.get(first, 0) + 1
    return columns


def _band_groups(sheet, row: int, max_col: int) -> Dict[int, str]:
    """グループ見出しの行を読み、結合を考慮して列ごとの所属を返す。"""
    if row < 1:
        return {}
    spans = {}
    for merged in sheet.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row:
            value = sheet.cell(row=merged.min_row, column=merged.min_col).value
            if isinstance(value, str) and value.strip():
                for col in range(merged.min_col, merged.max_col + 1):
                    spans[col] = value.strip()
    for col in range(1, max_col + 1):
        if col in spans:
            continue
        value = sheet.cell(row=row, column=col).value
        if isinstance(value, str) and value.strip():
            spans[col] = value.strip()
    return spans


def _read_title(sheet, header_row: int) -> str:
    """見出しより上にある最初の文字列をタイトルとみなす。"""
    for row in range(1, header_row):
        for col in range(1, 6):
            value = sheet.cell(row=row, column=col).value
            if isinstance(value, str) and value.strip():
                return value.strip()
    return ""


# ----------------------------------------------------------------------
def _read_rows(sheet, first_row: int, columns: Dict[str, int], language: str):
    """記入された行を読む。空行は飛ばす。"""
    rows: List[Row] = []
    warnings: List[str] = []
    group = subgroup = ""

    for index in range(first_row, (sheet.max_row or first_row) + 1):
        raw = {key: sheet.cell(row=index, column=col).value
               for key, col in columns.items()}
        if all(v in (None, "") for v in raw.values()):
            continue

        row = Row(row=index)
        row.name = _text(raw.get("name"))
        # 大項目・中項目は書かれた行だけに入るので、下の行へ引き継ぐ
        group = _text(raw.get("group")) or group
        subgroup = _text(raw.get("subgroup")) or subgroup
        row.group, row.subgroup = group, subgroup
        row.no = _text(raw.get("no"))
        row.predecessor = _text(raw.get("predecessor"))
        row.member = _text(raw.get("member"))
        row.status = _text(raw.get("status"))

        try:
            row.start = _date(raw.get("start"))
            row.end = _date(raw.get("end"))
            row.actual_start = _date(raw.get("actual_start"))
            row.actual_end = _date(raw.get("actual_end"))
            row.days = _int(raw.get("days"))
            row.actual_days = _int(raw.get("actual_days"))
            row.delay = _int(raw.get("delay"))
            row.progress = _ratio(raw.get("progress"))
            row.effort = _number(raw.get("effort"))
        except ValueError as exc:
            warnings.append(message(language, "row_prefix", row=index,
                                    reason=_reason(exc, language)))
            continue

        if not row.name and not row.has_plan:
            continue
        rows.append(row)
        if len(rows) > MAX_ROWS:
            warnings.append(message(language, "too_many_rows", maximum=MAX_ROWS))
            break

    return rows, warnings


# ----------------------------------------------------------------------
def _read_config(book) -> Dict[str, Any]:
    """設定シートから表示期間と稼働日を読む。無ければ空。"""
    sheet = next((book[name] for name in book.sheetnames if name in CONFIG_SHEETS), None)
    if sheet is None:
        return {}
    pairs: Dict[str, Any] = {}
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row or 1, 40),
                              max_col=3, values_only=True):
        label = row[1] if len(row) > 1 else None
        value = row[2] if len(row) > 2 else None
        if isinstance(label, str) and label.strip():
            pairs[label.strip()] = value

    def first(kind):
        for label in CONFIG_KEYS[kind]:
            if label in pairs:
                return pairs[label]
        return None

    config: Dict[str, Any] = {}
    start = _date_or_none(first("start"))
    if start:
        config["start"] = start
    period = first("period")
    if isinstance(period, (int, float)):
        config["period_days"] = int(period)
    unit = first("unit")
    if isinstance(unit, str) and unit.strip() in UNIT_LABELS:
        config["unit"] = UNIT_LABELS[unit.strip()]

    # 稼働曜日 (曜日名は言語ごとに違うので、まとめた表で引く)
    working = set()
    for label, value in pairs.items():
        index = WEEKDAY_LABELS.get(label)
        if index is not None and str(value).strip() in WORK_ON_LABELS:
            working.add(index)
    if working:
        config["workdays"] = [WEEKDAY_KEYS[i] for i in sorted(working)]

    holidays = []
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row or 1, 400),
                              min_col=5, max_col=5, values_only=True):
        day = _date_or_none(row[0])
        if day:
            holidays.append(day)
    config["holidays"] = holidays
    return config


def _read_members(book) -> List[str]:
    sheet = next((book[name] for name in book.sheetnames if name in MEMBER_SHEETS), None)
    if sheet is None:
        return []
    names = []
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row or 1, 200),
                               min_col=2, max_col=2, values_only=True):
        value = _text(row[0])
        if value and value not in MEMBER_SKIP:
            names.append(value)
    return names


def _build_spec(config, rows: List[Row], book, filename: str, language: str) -> BlankWBS:
    """表示に使う期間を決める。設定シートが無ければ記入内容から割り出す。"""
    days = [d for row in rows
            for d in (row.start, row.end, row.actual_start, row.actual_end) if d]

    start = config.get("start")
    period = config.get("period_days")
    if start is None:
        start = min(days) if days else _dt.date.today()
        start = start.replace(day=1)
    if period is None:
        last = max(days) if days else start
        period = max((last - start).days + 14, 30)

    title = Path(filename).stem if filename else message(language, "imported_title")
    spec = BlankWBS(
        title=title,
        start=start,
        period_days=min(max(period, 1), 3660),
        unit=config.get("unit") or UNIT_WEEK,
        rows=0,
        members=_read_members(book),
        workdays=config.get("workdays") or ["mon", "tue", "wed", "thu", "fri"],
        # 休日は設定シートの一覧をそのまま使う。無ければ祝日を補う。
        japanese_holidays=not config.get("holidays"),
        holidays=config.get("holidays") or [],
        language=language,
    )
    if spec.unit not in VALID_UNITS:
        spec.unit = UNIT_WEEK
    return spec


# ----------------------------------------------------------------------
# セルの値の変換
# ----------------------------------------------------------------------
class _CellError(str):
    """セルを読めなかったときの理由。文言は行を組み立てるときに解決する。

    ``str`` を継承しているので、そのまま例外に載せても壊れない。
    """

    def __new__(cls, key: str, value):
        self = super().__new__(cls, f"{key}:{value!r}")
        self.key = key
        self.value = value
        return self

def _reason(exc: Exception, language: str) -> str:
    """例外から、その言語での理由を組み立てる。"""
    detail = exc.args[0] if exc.args else ""
    if isinstance(detail, _CellError):
        return message(language, detail.key, value=detail.value)
    return str(detail)


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _date(value) -> Optional[_dt.date]:
    day = _date_or_none(value)
    if day is None and value not in (None, ""):
        raise ValueError(_CellError("cell_bad_date", value))
    return day


def _date_or_none(value) -> Optional[_dt.date]:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if isinstance(value, str) and value.strip():
        text = value.strip().replace("/", "-")
        for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%m-%d"):
            try:
                parsed = _dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
            if fmt == "%m-%d":
                parsed = parsed.replace(year=_dt.date.today().year)
            return parsed.date()
    return None


def _int(value) -> Optional[int]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip().replace("日", "").strip()
    if not text:
        return None
    try:
        return int(float(text))
    except ValueError:
        raise ValueError(_CellError("cell_bad_days", value))


def _number(value) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except ValueError:
        raise ValueError(_CellError("cell_bad_number", value))


def _ratio(value) -> Optional[float]:
    """進捗。``0.8`` でも ``80%`` でも ``80`` でも受ける。"""
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = str(value).strip()
        try:
            number = float(text.rstrip("%")) / (100.0 if text.endswith("%") else 1.0)
        except ValueError:
            raise ValueError(_CellError("cell_bad_progress", value))
    if number > 1.0:
        number /= 100.0
    return min(max(number, 0.0), 1.0)
