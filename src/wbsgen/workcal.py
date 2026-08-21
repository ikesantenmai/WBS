"""稼働日カレンダー。

空の WBS では日程の計算はしないが、日単位表示で休日の列を色分けし、
「設定」シートに休日一覧を載せるために、どの日が休みかは判定する。
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Iterable, Sequence

WEEKDAY_KEYS = ("mon", "tue", "wed", "thu", "fri", "sat", "sun")
WEEKDAY_JP = ("月", "火", "水", "木", "金", "土", "日")

#: 既定の稼働曜日 (月〜金)
DEFAULT_WORKDAYS = ("mon", "tue", "wed", "thu", "fri")


def parse_weekdays(values: Iterable[str]) -> frozenset:
    """``["mon", "tue"]`` / ``["月", "火"]`` を weekday 番号の集合に変換する。"""
    out = set()
    for value in values:
        key = str(value).strip().lower()
        if key in WEEKDAY_KEYS:
            out.add(WEEKDAY_KEYS.index(key))
        elif value in WEEKDAY_JP:
            out.add(WEEKDAY_JP.index(value))
        else:
            raise ValueError(f"曜日の指定が不正です: {value!r}")
    return frozenset(out)


@dataclass(frozen=True)
class WorkCalendar:
    """稼働日判定を行うカレンダー。

    Attributes:
        workdays: 稼働扱いにする曜日 (0=月 .. 6=日)。
        holidays: 休日にする日付 (祝日・全社休業日など)。
        extra_workdays: ``workdays`` に含まれない曜日でも稼働扱いにする日付。
    """

    workdays: frozenset = field(default_factory=lambda: parse_weekdays(DEFAULT_WORKDAYS))
    holidays: frozenset = frozenset()
    extra_workdays: frozenset = frozenset()

    @classmethod
    def build(
        cls,
        workdays: Sequence[str] = DEFAULT_WORKDAYS,
        holidays: Iterable[_dt.date] = (),
        extra_workdays: Iterable[_dt.date] = (),
    ) -> "WorkCalendar":
        return cls(
            workdays=parse_weekdays(workdays),
            holidays=frozenset(holidays),
            extra_workdays=frozenset(extra_workdays),
        )

    def is_workday(self, day: _dt.date) -> bool:
        if day in self.extra_workdays:
            return True
        if day in self.holidays:
            return False
        return day.weekday() in self.workdays

    def is_holiday(self, day: _dt.date) -> bool:
        return not self.is_workday(day)

    def holidays_between(self, start: _dt.date, end: _dt.date) -> "list[_dt.date]":
        """期間内の休日 (設定シートに載せる一覧)。"""
        out = []
        day = start
        step = _dt.timedelta(days=1)
        while day <= end:
            if day in self.holidays:
                out.append(day)
            day += step
        return out


# ----------------------------------------------------------------------
# 日本の祝日
# ----------------------------------------------------------------------
def _vernal_equinox(year: int) -> _dt.date:
    day = int(20.8431 + 0.242194 * (year - 1980) - (year - 1980) // 4)
    return _dt.date(year, 3, day)


def _autumnal_equinox(year: int) -> _dt.date:
    day = int(23.2488 + 0.242194 * (year - 1980) - (year - 1980) // 4)
    return _dt.date(year, 9, day)


def _nth_weekday(year: int, month: int, weekday: int, nth: int) -> _dt.date:
    first = _dt.date(year, month, 1)
    offset = (weekday - first.weekday()) % 7
    return first + _dt.timedelta(days=offset + 7 * (nth - 1))


def japanese_holidays(year: int) -> "list[_dt.date]":
    """``year`` 年の日本の国民の祝日 (振替休日・国民の休日を含む)。

    2007 年以降の祝日法に対応する。山の日は 2016 年から、
    天皇誕生日は 2020 年から 2/23 として扱う。
    """
    fixed = [
        (1, 1),    # 元日
        (2, 11),   # 建国記念の日
        (4, 29),   # 昭和の日
        (5, 3),    # 憲法記念日
        (5, 4),    # みどりの日
        (5, 5),    # こどもの日
        (11, 3),   # 文化の日
        (11, 23),  # 勤労感謝の日
    ]
    days = [_dt.date(year, m, d) for m, d in fixed]
    days.append(_nth_weekday(year, 1, 0, 2))   # 成人の日
    days.append(_nth_weekday(year, 7, 0, 3))   # 海の日
    days.append(_nth_weekday(year, 9, 0, 3))   # 敬老の日
    days.append(_nth_weekday(year, 10, 0, 2))  # スポーツの日
    days.append(_vernal_equinox(year))         # 春分の日
    days.append(_autumnal_equinox(year))       # 秋分の日
    if year >= 2016:
        days.append(_dt.date(year, 8, 11))     # 山の日
    if year >= 2020:
        days.append(_dt.date(year, 2, 23))     # 天皇誕生日
    else:
        days.append(_dt.date(year, 12, 23))

    known = set(days)
    # 国民の休日 (祝日に挟まれた平日)
    for day in sorted(known):
        nxt = day + _dt.timedelta(days=2)
        between = day + _dt.timedelta(days=1)
        if nxt in known and between not in known and between.weekday() != 6:
            days.append(between)

    # 振替休日 (日曜と重なった祝日の直後の平日)
    known = set(days)
    for day in sorted(known):
        if day.weekday() == 6:
            sub = day + _dt.timedelta(days=1)
            while sub in known:
                sub += _dt.timedelta(days=1)
            days.append(sub)

    return sorted(set(days))


def japanese_holidays_range(start: _dt.date, end: _dt.date) -> "list[_dt.date]":
    out = []
    for year in range(start.year, end.year + 1):
        out.extend(d for d in japanese_holidays(year) if start <= d <= end)
    return sorted(set(out))
