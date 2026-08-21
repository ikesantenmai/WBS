"""空の WBS の仕様。

期間・表示単位・記入用の空行数といった「用紙の寸法」だけを持つ。
タスクは扱わない。
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Any, Dict, Optional, Sequence

from .timeline import UNIT_WEEK, VALID_UNITS
from .workcal import (
    DEFAULT_WORKDAYS,
    WorkCalendar,
    japanese_holidays_range,
    parse_weekdays,
)

DEFAULT_ROWS = 40
DEFAULT_MONTHS = 12
MAX_ROWS = 2000
MAX_PERIOD_DAYS = 3660      # 10 年
MAX_MEMBERS = 100


class SpecError(ValueError):
    """指定の不備。利用者に見せるメッセージを持つ。"""


@dataclass
class BlankWBS:
    """空の WBS 1 枚ぶんの仕様。"""

    title: str = "プロジェクト スケジュール"
    start: _dt.date = field(default_factory=_dt.date.today)
    period_days: int = 365
    unit: str = UNIT_WEEK
    rows: int = DEFAULT_ROWS
    members: list = field(default_factory=list)
    workdays: list = field(default_factory=lambda: list(DEFAULT_WORKDAYS))
    japanese_holidays: bool = True
    holidays: list = field(default_factory=list)

    @property
    def end(self) -> _dt.date:
        """表示期間の最終日。"""
        return self.start + _dt.timedelta(days=self.period_days - 1)

    def calendar(self) -> WorkCalendar:
        """休日判定に使うカレンダー。"""
        holidays = set(self.holidays)
        if self.japanese_holidays:
            pad = _dt.timedelta(days=40)
            holidays.update(japanese_holidays_range(self.start - pad, self.end + pad))
        return WorkCalendar.build(workdays=self.workdays, holidays=holidays)


# ----------------------------------------------------------------------
def add_months(day: _dt.date, months: int) -> _dt.date:
    """``day`` の ``months`` か月後 (月末は丸める)。"""
    total = day.month - 1 + months
    year = day.year + total // 12
    month = total % 12 + 1
    last = [31, 29 if _is_leap(year) else 28, 31, 30, 31, 30,
            31, 31, 30, 31, 30, 31][month - 1]
    return _dt.date(year, month, min(day.day, last))


def _is_leap(year: int) -> bool:
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


# ----------------------------------------------------------------------
def build(
    start: _dt.date,
    end: Optional[_dt.date] = None,
    period_days: Optional[int] = None,
    months: Optional[int] = None,
    unit: str = UNIT_WEEK,
    rows: int = DEFAULT_ROWS,
    title: Optional[str] = None,
    members: Optional[Sequence[str]] = None,
    workdays: Optional[Sequence[str]] = None,
    japanese_holidays: bool = True,
    holidays: Optional[Sequence[_dt.date]] = None,
) -> BlankWBS:
    """期間を指定して :class:`BlankWBS` を組み立てる。

    期間は ``end`` / ``period_days`` / ``months`` のどれかで指定する。
    いずれも省略した場合は 12 か月。
    """
    if not isinstance(start, _dt.date):
        raise SpecError("開始日を指定してください。")
    if unit not in VALID_UNITS:
        raise SpecError(f"表示単位が不正です: {unit!r} ({'/'.join(VALID_UNITS)})")

    if end is not None:
        if end < start:
            raise SpecError(f"終了日が開始日より前です: {start} 〜 {end}")
        days = (end - start).days + 1
    elif period_days is not None:
        days = int(period_days)
    else:
        days = (add_months(start, months or DEFAULT_MONTHS) - start).days

    if days < 1:
        raise SpecError(f"期間は 1 日以上で指定してください: {days}")
    if days > MAX_PERIOD_DAYS:
        raise SpecError(f"期間が長すぎます (上限 10 年): {days} 日")
    if not 0 <= rows <= MAX_ROWS:
        raise SpecError(f"行数は 0〜{MAX_ROWS} の範囲で指定してください: {rows}")

    names = [str(m).strip() for m in (members or []) if str(m).strip()]
    if len(names) > MAX_MEMBERS:
        raise SpecError(f"担当者は {MAX_MEMBERS} 人までです: {len(names)} 人")

    day_keys = list(workdays) if workdays else list(DEFAULT_WORKDAYS)
    try:
        parse_weekdays(day_keys)
    except ValueError as exc:
        raise SpecError(str(exc))

    return BlankWBS(
        title=(title or "").strip() or f"{start.year}年 スケジュール",
        start=start,
        period_days=days,
        unit=unit,
        rows=rows,
        members=names,
        workdays=day_keys,
        japanese_holidays=japanese_holidays,
        holidays=sorted(set(holidays or [])),
    )


# ----------------------------------------------------------------------
def to_dict(spec: BlankWBS) -> Dict[str, Any]:
    """Web API とのやり取り用。"""
    return {
        "title": spec.title,
        "start": spec.start.isoformat(),
        "period_days": spec.period_days,
        "end": spec.end.isoformat(),
        "unit": spec.unit,
        "rows": spec.rows,
        "members": list(spec.members),
        "workdays": list(spec.workdays),
        "japanese_holidays": spec.japanese_holidays,
        "holidays": [d.isoformat() for d in spec.holidays],
    }


def from_dict(data: Dict[str, Any]) -> BlankWBS:
    """辞書 (JSON) から組み立てる。日付は ``YYYY-MM-DD``。"""
    if not isinstance(data, dict):
        raise SpecError("指定はマッピングである必要があります。")
    return build(
        start=parse_date(data.get("start"), "start"),
        end=parse_date(data["end"], "end") if data.get("end") else None,
        period_days=_int(data.get("period_days"), "period_days"),
        months=_int(data.get("months"), "months"),
        unit=str(data.get("unit", UNIT_WEEK)),
        rows=_int(data.get("rows"), "rows") if data.get("rows") is not None else DEFAULT_ROWS,
        title=data.get("title"),
        members=data.get("members") or [],
        workdays=data.get("workdays") or None,
        japanese_holidays=bool(data.get("japanese_holidays", True)),
        holidays=[parse_date(d, "holidays") for d in (data.get("holidays") or [])],
    )


def parse_date(value, field_name: str) -> _dt.date:
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    if not value:
        raise SpecError(f"{field_name} を指定してください (YYYY-MM-DD)。")
    try:
        return _dt.datetime.strptime(str(value).strip().replace("/", "-"), "%Y-%m-%d").date()
    except ValueError:
        raise SpecError(f"{field_name} は YYYY-MM-DD 形式で指定してください: {value!r}")


def _int(value, field_name: str) -> Optional[int]:
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        raise SpecError(f"{field_name} は整数で指定してください: {value!r}")
