"""本日の状況（日次チェック）。

読み込んだ WBS を基準日から見て、その日に手を打つべきことを拾う。

- 本日開始予定 / 本日終了予定
- 遅延タスク
- 未着手タスク
- スケジュール整合性チェック（書き方の食い違い）
- 本日アサインがない担当者

行そのものは返さず **行番号** (:attr:`wbsgen.importer.Row.row`) を返す。
画面はすでに全行を持っているので、番号で引けば同じ内容を二度送らずに済む。

遅れと状態は :func:`wbsgen.importer.resolve` が基準日から求めているので、
ここに渡す行は同じ基準日で解決済みであること。
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from typing import List, Optional

from .workload import member_names, split_names

#: 先行の欄の区切り (「1, 2」「3・4」などをばらす)
PREDECESSORS = re.compile(r"[/／|｜、，,;；・･\s]+")


@dataclass
class Check:
    """整合性チェック 1 つぶんの結果。``rows`` が空なら問題なし。"""

    kind: str
    rows: List[int] = field(default_factory=list)


@dataclass
class Digest:
    """基準日から見た、その日の状況。"""

    date: _dt.date
    starting: List[int] = field(default_factory=list)
    ending: List[int] = field(default_factory=list)
    delayed: List[int] = field(default_factory=list)
    not_started: List[int] = field(default_factory=list)
    checks: List[Check] = field(default_factory=list)
    #: 担当欄から拾えた名前 (アサイン無しを数える母数)
    members: List[str] = field(default_factory=list)
    #: そのうち、基準日にかかる予定が 1 つも無い人
    idle_members: List[str] = field(default_factory=list)


def build(rows, calendar=None, base_date: Optional[_dt.date] = None) -> Digest:
    """行の一覧から、基準日の状況をまとめる。

    読み込みの時点で空行は落ちているので、来た行はすべて数える。
    大項目だけの行 (まとめの行) も、遅れていれば遅れとして拾う
    (画面の上に出ている集計と件数がそろうように)。
    """
    base = base_date or _dt.date.today()
    tasks = list(rows)

    digest = Digest(date=base)
    digest.starting = [row.row for row in tasks if row.start == base]
    digest.ending = [row.row for row in tasks if row.end == base]
    digest.delayed = [row.row for row in sorted(
        (row for row in tasks if row.delay),
        key=lambda row: (-(row.delay or 0), row.row))]
    digest.not_started = [row.row for row in sorted(
        (row for row in tasks
         if not row.has_actual_record and not row.is_finished
         and (row.start or row.end)),
        key=lambda row: (row.start or _dt.date.max, row.row))]
    digest.checks = _checks(tasks)
    digest.members = member_names(tasks)
    digest.idle_members = _idle_members(tasks, digest.members, base)
    return digest


# ---------------------------------------------------------------- 整合性
def _checks(tasks) -> List[Check]:
    """順に確かめる。問題が無くても「確かめた」ことが判るよう、必ず返す。"""
    return [
        Check("end_before_start", _rows(tasks, _end_before_start)),
        Check("actual_end_before_start", _rows(tasks, _actual_end_before_start)),
        Check("plan_incomplete", _rows(tasks, _plan_incomplete)),
        Check("actual_end_without_start", _rows(tasks, _actual_end_without_start)),
        Check("progress_without_actual", _rows(tasks, _progress_without_actual)),
        Check("done_without_actual_end", _rows(tasks, _done_without_actual_end)),
        Check("no_dates", _rows(tasks, _no_dates)),
        Check("days_rewritten", _rows(tasks, _days_rewritten)),
        Check("actual_days_rewritten", _rows(tasks, _actual_days_rewritten)),
        Check("predecessor_missing", _predecessor_missing(tasks)),
        Check("predecessor_order", _predecessor_order(tasks)),
    ]


def _rows(tasks, test) -> List[int]:
    return [row.row for row in tasks if test(row)]


def _end_before_start(row) -> bool:
    """予定: 終了日が開始日より前。"""
    return bool(row.start and row.end and row.end < row.start)


def _actual_end_before_start(row) -> bool:
    """実績: 終了日が開始日より前。"""
    return bool(row.actual_start and row.actual_end
                and row.actual_end < row.actual_start)


def _plan_incomplete(row) -> bool:
    """予定の開始日と終了日が片方しか無い (日数からも決まらない)。"""
    return bool(row.start) != bool(row.end)


def _actual_end_without_start(row) -> bool:
    """実績の終了日はあるのに、開始日が無い。"""
    return bool(row.actual_end and not row.actual_start)


def _progress_without_actual(row) -> bool:
    """進捗が入っているのに、実績の開始日が無い。"""
    return bool(row.progress and not row.actual_start)


def _done_without_actual_end(row) -> bool:
    """進捗が 100% なのに、実績の終了日が無い。"""
    return row.progress == 1.0 and row.actual_end is None


def _no_dates(row) -> bool:
    """何か書かれている行なのに、予定も実績も 1 つも入っていない。"""
    return not (row.start or row.end or row.has_actual_record)


def _days_rewritten(row) -> bool:
    """書かれていた予定日数が、開始日・終了日から数えた日数と違う。"""
    return "days" in row.derived and row.written.get("days") is not None


def _actual_days_rewritten(row) -> bool:
    """書かれていた実績日数が、実績の開始日・終了日と合わない。"""
    return "actual_days" in row.derived and row.written.get("actual_days") is not None


def _predecessor_missing(tasks) -> List[int]:
    """先行に書かれた項番が、この表に見つからない。"""
    known = {str(row.no).strip() for row in tasks if str(row.no).strip()}
    return [row.row for row in tasks
            if any(no not in known for no in _predecessors(row))]


def _predecessor_order(tasks) -> List[int]:
    """先行タスクの予定終了日より前に始まる予定になっている。"""
    by_no = {}
    for row in tasks:
        key = str(row.no).strip()
        if key:
            by_no.setdefault(key, row)

    out = []
    for row in tasks:
        if not row.start:
            continue
        for no in _predecessors(row):
            before = by_no.get(no)
            if before is not None and before.end and before.end > row.start:
                out.append(row.row)
                break
    return out


def _predecessors(row) -> List[str]:
    """先行の欄に書かれている項番。"""
    return [piece for piece in PREDECESSORS.split(str(row.predecessor or "").strip())
            if piece]


# ---------------------------------------------------------------- 担当
def _idle_members(tasks, names: List[str], base: _dt.date) -> List[str]:
    """基準日にかかる**予定**が 1 つも無い人を、名前の並び順で返す。

    数え方は要員稼働チェックと同じ。担当欄は「佐々木（高瀬）」のように
    他の名前を含むので、書かれた文字列に名前が含まれるかで数える。
    完了しているかどうかは見ない (予定の入り方をそろえて見るため)。
    """
    today = [str(row.member) for row in tasks
             if row.member and row.start and row.end
             and row.start <= base <= row.end]
    return [name for name in names
            if not any(name in member for member in today)]


__all__ = ["Check", "Digest", "build", "member_names", "split_names"]
