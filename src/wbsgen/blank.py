"""空の WBS の組み立て。

期間だけを決めて、記入用の空行が並んだ WBS を作る。
「まず枠を配って各自に埋めてもらう」という使い方のための入口。
"""

from __future__ import annotations

import datetime as _dt
from typing import Optional

from .loader import MAX_BLANK_ROWS, ProjectError
from .model import UNIT_WEEK, CalendarConfig, ChartConfig, Member, Project

DEFAULT_ROWS = 40
DEFAULT_MONTHS = 12


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


def build_blank(
    start: _dt.date,
    end: Optional[_dt.date] = None,
    period_days: Optional[int] = None,
    months: Optional[int] = None,
    unit: str = UNIT_WEEK,
    rows: int = DEFAULT_ROWS,
    title: str = "プロジェクト スケジュール",
    base_date: Optional[_dt.date] = None,
    members: Optional[list] = None,
) -> Project:
    """期間を指定して、中身が空の :class:`~wbsgen.model.Project` を返す。

    期間は ``end`` / ``period_days`` / ``months`` のどれかで指定する。
    いずれも省略した場合は 12 か月。
    """
    if end is not None:
        if end < start:
            raise ProjectError(f"終了日が開始日より前です: {start} 〜 {end}")
        days = (end - start).days + 1
    elif period_days is not None:
        days = int(period_days)
    else:
        days = (add_months(start, months or DEFAULT_MONTHS) - start).days

    if days < 1:
        raise ProjectError(f"期間は 1 日以上で指定してください: {days}")
    if days > 3660:
        raise ProjectError(f"期間が長すぎます (上限 10 年): {days} 日")
    if not 0 <= rows <= MAX_BLANK_ROWS:
        raise ProjectError(f"行数は 0〜{MAX_BLANK_ROWS} の範囲で指定してください: {rows}")

    return Project(
        title=title,
        chart=ChartConfig(
            start=start,
            period_days=days,
            unit=unit,
            base_date=base_date or _dt.date.today(),
        ),
        calendar=CalendarConfig(),
        members=[Member(name=name) for name in (members or [])],
        tasks=[],
        blank_rows=rows,
    )
