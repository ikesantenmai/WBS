"""チャート横軸 (日程表) の構築。

表示単位 (日/週/月) に応じて列を割り当て、任意の日付を
「列インデックス + 列内比率」に写像する。

見出しは元ファイルと同じ 2 段構成にする。

===== ====================== ==========================
単位   上段 (行 3)             下段 (行 4)
===== ====================== ==========================
日     月 (``m"月"``)          日 (``d``) ＋ 曜日
週     月 (``m"月"``)          週の開始日 (``m/d``)
月     年 (``yyyy"年"``)       月 (``m"月"``)
===== ====================== ==========================

上段は区切りが変わる列にだけ値を置く (元ファイルはセルを結合していない)。

週の列は**チャート表示開始日から 7 日刻み**で並ぶ。曜日には合わせない
(元ファイルは 2026-02-01 の日曜始まりで 2/1, 2/8, 2/15 … と並んでいる)。
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from typing import List

from .workcal import WorkCalendar

#: 表示単位
UNIT_DAY = "day"
UNIT_WEEK = "week"
UNIT_MONTH = "month"
VALID_UNITS = (UNIT_DAY, UNIT_WEEK, UNIT_MONTH)

WEEKDAY_JP = ("月", "火", "水", "木", "金", "土", "日")

#: 単位ごとの (上段の数値書式, 下段の数値書式)
HEADER_FORMATS = {
    UNIT_DAY: ('m"月"', "d"),
    UNIT_WEEK: ('m"月"', "m/d"),
    UNIT_MONTH: ('yyyy"年"', 'm"月"'),
}


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
            # 表示開始日を起点に 7 日刻み (曜日には合わせない)
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

    # ------------------------------------------------------------------
    # 見出し
    # ------------------------------------------------------------------
    @property
    def formats(self):
        """(上段の数値書式, 下段の数値書式)。"""
        return HEADER_FORMATS[self.unit]

    def _group_key(self, col: Column):
        """上段の区切り。日/週は月ごと、月表示は年ごと。"""
        if self.unit == UNIT_MONTH:
            return col.start.year
        return (col.start.year, col.start.month)

    def header_top(self):
        """上段の見出し。区切りが変わる列にだけ値を置く。

        ``(列インデックス, 列数, 日付, 表示文字列)`` の列を返す。
        """
        if not self.columns:
            return []
        out = []
        current = None
        first = 0
        for col in self.columns:
            key = self._group_key(col)
            if key != current:
                if current is not None:
                    out.append(self._top_entry(first, col.index - first))
                current = key
                first = col.index
        out.append(self._top_entry(first, len(self.columns) - first))
        return out

    def _top_entry(self, first: int, span: int):
        day = self.columns[first].start
        text = f"{day.year}年" if self.unit == UNIT_MONTH else f"{day.month}月"
        return (first, span, day, text)

    def header_bottom(self):
        """下段の見出し。全ての列に値を置く。"""
        return [
            (col.index, col.start, self.column_label(col))
            for col in self.columns
        ]

    def column_label(self, col: Column) -> str:
        """下段の表示文字列。"""
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
