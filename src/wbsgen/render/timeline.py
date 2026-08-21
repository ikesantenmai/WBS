"""チャート横軸 (日程表) の構築。

表示単位 (日/週/月) に応じて列を割り当て、任意の日付を
「列インデックス + 列内比率」に写像する。
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from typing import List

from ..model import UNIT_DAY, UNIT_MONTH, UNIT_WEEK
from ..workcal import WorkCalendar

WEEKDAY_JP = ("月", "火", "水", "木", "金", "土", "日")


@dataclass
class Column:
    """チャートの 1 列。"""

    index: int          # 0 始まりの列番号
    start: _dt.date     # この列が表す期間の開始日
    end: _dt.date       # 期間の終了日 (含む)

    @property
    def days(self) -> int:
        return (self.end - self.start).days + 1


class Timeline:
    """チャートの横軸。"""

    def __init__(self, start: _dt.date, period_days: int, unit: str, calendar: WorkCalendar):
        self.unit = unit
        self.calendar = calendar
        self.start = start
        self.end = start + _dt.timedelta(days=max(period_days, 1) - 1)
        self.columns: List[Column] = self._build()

    # ------------------------------------------------------------------
    def _build(self) -> List[Column]:
        cols: List[Column] = []
        day = self.start
        i = 0
        if self.unit == UNIT_DAY:
            while day <= self.end:
                cols.append(Column(i, day, day))
                day += _dt.timedelta(days=1)
                i += 1
        elif self.unit == UNIT_WEEK:
            # 週の開始は月曜に揃える
            day -= _dt.timedelta(days=day.weekday())
            self.start = day
            while day <= self.end:
                cols.append(Column(i, day, day + _dt.timedelta(days=6)))
                day += _dt.timedelta(days=7)
                i += 1
        else:  # month
            day = day.replace(day=1)
            self.start = day
            while day <= self.end:
                nxt = (day.replace(day=28) + _dt.timedelta(days=4)).replace(day=1)
                cols.append(Column(i, day, nxt - _dt.timedelta(days=1)))
                day = nxt
                i += 1
        if cols:
            self.end = cols[-1].end
        return cols

    # ------------------------------------------------------------------
    def __len__(self) -> int:
        return len(self.columns)

    def position(self, day: _dt.date, end_of_day: bool = False) -> float:
        """日付を「列単位の連続座標」に写像する。

        戻り値 ``2.5`` は 3 列目の中央を意味する。範囲外はクランプする。
        ``end_of_day`` が真なら、その日の終端 (翌日の始点) を返す。
        """
        if not self.columns:
            return 0.0
        if day < self.start:
            return 0.0
        if day > self.end:
            return float(len(self.columns))
        for col in self.columns:
            if col.start <= day <= col.end:
                offset = (day - col.start).days + (1 if end_of_day else 0)
                return col.index + offset / col.days
        return float(len(self.columns))

    def clamp(self, day: _dt.date) -> _dt.date:
        return min(max(day, self.start), self.end)

    def in_range(self, start: _dt.date, end: _dt.date) -> bool:
        return not (end < self.start or start > self.end)

    # ------------------------------------------------------------------
    def band_labels(self):
        """上段ラベル (年月など) を ``(開始列, 列数, 表示文字列)`` で返す。"""
        if not self.columns:
            return []
        out = []
        current = None
        first = 0
        for col in self.columns:
            key = (col.start.year, col.start.month) if self.unit != UNIT_MONTH else col.start.year
            if key != current:
                if current is not None:
                    out.append((first, col.index - first, self._band_text(current)))
                current = key
                first = col.index
        out.append((first, len(self.columns) - first, self._band_text(current)))
        return out

    def _band_text(self, key) -> str:
        if self.unit == UNIT_MONTH:
            return f"{key}年"
        year, month = key
        return f"{year}/{month}"

    def column_label(self, col: Column) -> str:
        """下段ラベル。"""
        if self.unit == UNIT_DAY:
            return str(col.start.day)
        if self.unit == UNIT_WEEK:
            return f"{col.start.month}/{col.start.day}"
        return f"{col.start.month}月"

    def weekday_label(self, col: Column) -> str:
        """日単位表示のときの曜日ラベル。"""
        if self.unit != UNIT_DAY:
            return ""
        return WEEKDAY_JP[col.start.weekday()]

    def is_rest_column(self, col: Column) -> bool:
        """列全体が非稼働日か (日単位のみ意味を持つ)。"""
        if self.unit != UNIT_DAY:
            return False
        return not self.calendar.is_workday(col.start)

    def is_saturday_column(self, col: Column) -> bool:
        return self.unit == UNIT_DAY and col.start.weekday() == 5
