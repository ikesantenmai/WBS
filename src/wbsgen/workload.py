"""要員稼働チェックの組み立て。

「スケジュール」の担当欄に名前がある行の**予定期間**を日別に数え、
月ごとに「誰がどの日に何件のタスクを持っているか」を並べる。
稼働日なのに 0 件の日は空きとして拾う。

Excel への書き出しは :mod:`wbsgen.workbook` が行う。ここは数え方だけを
持つので、単体で確かめられる。
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import List, Optional

#: 担当欄の区切り。「吉田/菊池」「長見、本家」「佐々木（高瀬）」を分ける。
SEPARATORS = re.compile(r"[/／|｜、，,;；・･\s]+|[（(]|[）)]")

#: 名前の後ろに付きがちな敬称
HONORIFICS = ("さん", "サン", "様", "氏")

#: これより長いものは名前ではなく文章とみなす
MAX_NAME_LENGTH = 20

#: 作る月の上限 (長すぎる計画でシートが増えすぎないように)
MAX_MONTHS = 24

#: 並べる人数の上限
MAX_MEMBERS = 100


@dataclass
class MemberLoad:
    """1 人ぶんの、その月の稼働。"""

    name: str
    #: 日ごとのタスク数 (:attr:`WorkloadMonth.days` と同じ並び)
    counts: List[int] = field(default_factory=list)
    #: 稼働日のうち、タスクがある日の数 (休みの日は数えない)
    busy_days: int = 0
    #: 稼働日なのにタスクが無い日
    free_days: List[_dt.date] = field(default_factory=list)


@dataclass
class WorkloadMonth:
    """1 か月ぶんの稼働チェック。"""

    year: int
    month: int
    days: List[_dt.date] = field(default_factory=list)
    #: 各日が稼働日か (休日・休業日なら ``False``)
    workdays: List[bool] = field(default_factory=list)
    members: List[MemberLoad] = field(default_factory=list)

    @property
    def workday_count(self) -> int:
        return sum(1 for on in self.workdays if on)

    @property
    def start(self) -> _dt.date:
        return self.days[0]

    @property
    def end(self) -> _dt.date:
        return self.days[-1]


# ----------------------------------------------------------------------
def member_names(rows) -> List[str]:
    """担当欄に書かれている名前を、出てきた順に並べる。

    「吉田/菊池/虎岩」「長見、本家、池田」「佐々木（高瀬）」のように
    1 つのセルに複数書いてあるので、区切って 1 人ずつにする。敬称は落とす。
    """
    names: List[str] = []
    seen = set()
    for row in rows:
        for name in split_names(getattr(row, "member", "")):
            if name not in seen:
                seen.add(name)
                names.append(name)
                if len(names) >= MAX_MEMBERS:
                    return names
    return names


def split_names(text: str) -> List[str]:
    """担当欄 1 つぶんを、名前の並びに分ける。"""
    out = []
    for piece in SEPARATORS.split(str(text or "")):
        name = piece.strip()
        for honorific in HONORIFICS:
            if len(name) > len(honorific) and name.endswith(honorific):
                name = name[: -len(honorific)]
                break
        if name and len(name) <= MAX_NAME_LENGTH:
            out.append(name)
    return out


def build(rows, calendar, base_date: Optional[_dt.date] = None) -> List[WorkloadMonth]:
    """予定の入っている行から、月ごとの稼働チェックを組み立てる。

    予定の開始日と終了日が揃っている行だけを数える (期間が決まらないため)。
    担当が誰も書かれていなければ、空の一覧を返す。
    """
    spans = [(row.start, row.end) for row in rows
             if row.start and row.end and row.end >= row.start]
    tasks = [(row.start, row.end, str(row.member))
             for row in rows
             if row.start and row.end and row.end >= row.start and row.member]
    names = member_names(rows)
    if not spans or not names:
        return []

    # 計画のある期間はすべて並べる。担当の無い行はどの人にも数えないが、
    # その月が抜けると「誰も入っていない月」が見えなくなるため。
    first = min(start for start, _ in spans)
    last = max(end for _, end in spans)

    months = []
    for year, month in _month_range(first, last):
        days = _days_of(year, month, first, last)
        entry = WorkloadMonth(year=year, month=month, days=days,
                              workdays=[calendar.is_workday(d) for d in days])
        for name in names:
            entry.members.append(_load(name, entry, tasks))
        months.append(entry)
    return months


# ----------------------------------------------------------------------
def _month_range(first: _dt.date, last: _dt.date):
    """``first`` から ``last`` までにかかる (年, 月) を順に返す。"""
    year, month = first.year, first.month
    out = []
    while (year, month) <= (last.year, last.month) and len(out) < MAX_MONTHS:
        out.append((year, month))
        year, month = (year + 1, 1) if month == 12 else (year, month + 1)
    return out


def _days_of(year: int, month: int, first: _dt.date, last: _dt.date) -> List[_dt.date]:
    """その月の日付。計画の範囲からはみ出すぶんは切り落とす。"""
    start = max(_dt.date(year, month, 1), first)
    end = min(_last_day(year, month), last)
    return [start + _dt.timedelta(days=i) for i in range((end - start).days + 1)]


def _last_day(year: int, month: int) -> _dt.date:
    return (_dt.date(year + 1, 1, 1) if month == 12
            else _dt.date(year, month + 1, 1)) - _dt.timedelta(days=1)


def _load(name: str, entry: WorkloadMonth, tasks) -> MemberLoad:
    """1 人ぶんの日別のタスク数を数える。

    担当欄は「佐々木（高瀬）」のように他の名前を含むことがあるので、
    書かれた文字列に名前が**含まれるか**で数える (元の表と同じ数え方)。
    """
    mine = [(start, end) for start, end, member in tasks if name in member]
    load = MemberLoad(name=name)
    for day, working in zip(entry.days, entry.workdays):
        count = sum(1 for start, end in mine if start <= day <= end)
        load.counts.append(count)
        if not working:
            continue
        if count:
            load.busy_days += 1
        else:
            load.free_days.append(day)
    return load
