"""日程・遅れ・状態の導出。

元の Excel ツールと同じ計算規則を実装する。

- 予定終了日 = 予定開始日を 1 日目とした ``日数`` 稼働日目
- 実績日数 (未完了) = 実績開始日から現在日までの稼働日数
- 遅れ = 予定終了日から現在日までの稼働日数 - 1
        (予定終了日が現在日より前で、実績終了日が無い場合)
        未着手なら予定開始日を起点に同じ式で「開始遅れ」を求める
- 状態 = 完了 / 遅れ n 日 / 残り n 日 / 実行中 / あと n 日 / -
"""

from __future__ import annotations

import datetime as _dt
from typing import Optional

from .model import (
    KIND_MILESTONE,
    KIND_SUMMARY,
    CalendarConfig,
    Project,
    Task,
)
from .workcal import WorkCalendar, japanese_holidays_range

#: 状態表示の文言 (元ツールの「設定」シートの既定値)
STATUS_NONE = "-"
STATUS_DONE = "完了"
STATUS_RUNNING = "実行中"
STATUS_REMAIN = "残り %d 日"
STATUS_DELAY = "遅れ %d 日"
STATUS_NEAR = "あと %d 日"


def build_calendar(cfg: CalendarConfig, start: _dt.date, end: _dt.date) -> WorkCalendar:
    """設定とプロジェクト期間から稼働日カレンダーを組み立てる。"""
    holidays = set(cfg.holidays)
    if cfg.japanese_holidays:
        pad = _dt.timedelta(days=400)
        holidays.update(japanese_holidays_range(start - pad, end + pad))
    return WorkCalendar.build(
        workdays=cfg.workdays,
        holidays=holidays,
        extra_workdays=cfg.extra_workdays,
    )


def member_calendar(project: Project, base: WorkCalendar, member_name: str) -> WorkCalendar:
    """担当者固有の休日設定があればそれを反映したカレンダーを返す。"""
    member = project.member(member_name)
    if member is None:
        return base
    if not (member.workdays or member.holidays or member.extra_workdays):
        return base
    from .workcal import parse_weekdays

    return WorkCalendar(
        workdays=parse_weekdays(member.workdays) if member.workdays else base.workdays,
        holidays=base.holidays | frozenset(member.holidays),
        extra_workdays=base.extra_workdays | frozenset(member.extra_workdays),
    )


# ----------------------------------------------------------------------
def resolve_task(
    task: Task,
    cal: WorkCalendar,
    base_date: _dt.date,
    thresholds,
) -> Task:
    """1 タスクの終了日・遅れ・状態を導出して ``task`` を更新する。"""

    # --- 予定 ---------------------------------------------------------
    if task.start is not None:
        if task.end is None and task.days is not None:
            task.end = cal.end_date(task.start, task.days)
        elif task.days is None and task.end is not None:
            task.days = cal.workdays_between(task.start, task.end)
    if task.kind == KIND_MILESTONE and task.start is not None:
        task.end = task.start
        task.days = task.days or 1

    # --- 実績 ---------------------------------------------------------
    if task.actual_start is not None:
        if task.is_complete:
            if task.actual_end is None:
                days = task.actual_days or 1
                task.actual_end = cal.end_date(task.actual_start, days)
            if task.actual_days is None:
                task.actual_days = cal.workdays_between(task.actual_start, task.actual_end)
        else:
            # 未完了: 実績終了日は置かず、経過稼働日数を実績日数とする
            task.actual_end = None
            if task.actual_days is None:
                task.actual_days = cal.workdays_between(task.actual_start, base_date)
    elif task.actual_end is not None and task.actual_days:
        # 終了日と日数だけ与えられた場合は開始日を逆算しない (そのまま扱う)
        pass

    # --- 遅れ ---------------------------------------------------------
    task.delay = _delay_days(task, cal, base_date)

    # --- 状態 ---------------------------------------------------------
    task.status = _status_text(task, cal, base_date, thresholds)
    return task


def _delay_days(task: Task, cal: WorkCalendar, base_date: _dt.date) -> Optional[int]:
    if task.kind == KIND_SUMMARY or not task.has_plan:
        return None
    # 開始遅れ: 予定開始日を過ぎているのに未着手
    if task.actual_start is None and task.start < base_date:
        delay = cal.workdays_between(task.start, base_date) - 1
        if delay > 0:
            return delay
    # 終了遅れ: 予定終了日を過ぎているのに未完了
    if task.actual_end is None and task.end < base_date:
        delay = cal.workdays_between(task.end, base_date) - 1
        if delay > 0:
            return delay
    return None


def _status_text(task: Task, cal: WorkCalendar, base_date: _dt.date, thresholds) -> str:
    if task.kind == KIND_SUMMARY or not task.has_plan:
        return STATUS_NONE
    if task.actual_end is not None:
        return STATUS_DONE
    if task.delay:
        return STATUS_DELAY % task.delay
    if task.actual_start is not None:
        remain = cal.workdays_between(base_date, task.end) - 1
        if remain <= thresholds.exec_remain_days:
            return STATUS_REMAIN % max(remain, 0)
        return STATUS_RUNNING
    remain = cal.workdays_between(base_date, task.start) - 1
    if remain <= thresholds.start_near_days:
        return STATUS_NEAR % max(remain, 0)
    return STATUS_NONE



# ----------------------------------------------------------------------
def autoschedule(project: Project, cal: WorkCalendar) -> None:
    """開始日が未指定で先行タスクを持つ行に、開始日を割り当てる。

    先行タスクの予定終了日の翌稼働日を開始日とする (FS 関係)。
    先行が未解決の間は繰り返し、循環参照は :class:`ValueError` にする。
    """
    by_no = {str(t.no): t for t in project.tasks if t.no not in ("", None)}
    pending = [
        t for t in project.tasks
        if t.start is None and t.predecessor and t.days is not None
    ]
    while pending:
        progressed = []
        for task in pending:
            preds = _predecessors(task, by_no)
            if preds is None:
                progressed.append(task)   # 存在しない先行番号は無視する
                continue
            ends = []
            for pred in preds:
                if pred.start is None:
                    ends = None
                    break
                pred_end = pred.end or cal.end_date(pred.start, pred.days or 1)
                ends.append(pred_end)
            if ends is None:
                continue
            if ends:
                task.start = cal.next_workday(max(ends) + _dt.timedelta(days=1))
                task.end = cal.end_date(task.start, task.days or 1)
            progressed.append(task)
        if not progressed:
            names = ", ".join(t.name for t in pending)
            raise ValueError(f"先行タスクの参照が循環しています: {names}")
        pending = [t for t in pending if t not in progressed]


def _predecessors(task: Task, by_no):
    """先行タスクの一覧。参照先が 1 つも見つからない場合は ``None``。"""
    keys = [k.strip() for k in str(task.predecessor).replace("、", ",").split(",") if k.strip()]
    found = [by_no[k] for k in keys if k in by_no]
    return found if found else None


# ----------------------------------------------------------------------
def resolve_project(project: Project) -> WorkCalendar:
    """プロジェクト全体を解決し、既定の稼働日カレンダーを返す。

    チャート表示開始日・期間が未指定の場合はタスクの範囲から自動決定する。
    """
    def anchors():
        return [t.start for t in project.tasks if t.start] + [
            t.actual_start for t in project.tasks if t.actual_start
        ]

    base_date = project.chart.base_date or _dt.date.today()
    dates = anchors()
    span_start = min(dates) if dates else base_date
    span_end = (max(dates) if dates else base_date) + _dt.timedelta(days=730)

    cal = build_calendar(project.calendar, span_start, span_end)

    # 先行タスクからの自動配置 (循環参照はここで検出される)
    autoschedule(project, cal)

    dates = anchors()
    if not dates:
        raise ValueError(
            "日付を持つタスクが 1 件もありません。"
            "少なくとも 1 件に start を指定するか、先行タスクの起点を作ってください。"
        )
    span_start = min(dates)

    for task in project.tasks:
        resolve_task(task, member_calendar(project, cal, task.member), base_date,
                     project.chart.thresholds)

    # サマリ (工程) 行は配下タスクの範囲を集計する
    _rollup_summaries(project, cal, base_date)

    ends = [t.end for t in project.tasks if t.end] + [
        t.actual_end for t in project.tasks if t.actual_end
    ]
    if project.chart.start is None:
        project.chart.start = _dt.date(span_start.year, span_start.month, 1)
    if not project.chart.period_days:
        last = max(ends) if ends else span_start
        project.chart.period_days = (last - project.chart.start).days + 30
    return cal


def _rollup_summaries(project: Project, cal: WorkCalendar, base_date: _dt.date) -> None:
    """``kind: summary`` の行に、同じ大項目の配下タスクの範囲/進捗を集計する。"""
    for summary in project.tasks:
        if summary.kind != KIND_SUMMARY:
            continue
        children = [
            t
            for t in project.tasks
            if t is not summary and t.kind != KIND_SUMMARY and t.group == summary.group
        ]
        starts = [t.start for t in children if t.start]
        ends = [t.end for t in children if t.end]
        if starts and summary.start is None:
            summary.start = min(starts)
        if ends and summary.end is None:
            summary.end = max(ends)
        if summary.start and summary.end and summary.days is None:
            summary.days = cal.workdays_between(summary.start, summary.end)
        if summary.effort is None:
            efforts = [t.effort for t in children if t.effort]
            summary.effort = sum(efforts) if efforts else None
        if summary.progress is None:
            weighted = [(t.progress or 0.0) * (t.days or 1) for t in children if t.has_plan]
            total = sum(t.days or 1 for t in children if t.has_plan)
            if total:
                summary.progress = round(sum(weighted) / total, 4)
        summary.status = STATUS_NONE
        summary.delay = None


def group_effort(project: Project, group: str) -> float:
    """大項目の工数合計 (人日)。"""
    return sum(
        t.effort or 0.0
        for t in project.tasks
        if t.group == group and t.kind != KIND_SUMMARY
    )


def inazuma_points(project: Project, cal: WorkCalendar, base_date: _dt.date):
    """イナズマ線の折れ点を ``(task, achieved_date)`` の列で返す。

    各タスクの進捗率を予定期間に当てはめ、「進捗上の到達日」を求める。
    完了タスクは予定終了日、未着手は予定開始日に落ちる。
    """
    points = []
    for task in project.tasks:
        if task.kind != "task" or not task.has_plan:
            continue
        progress = task.progress or 0.0
        if progress <= 0:
            achieved = task.start
        elif progress >= 1:
            achieved = task.end
        else:
            done = max(1, int(round((task.days or 1) * progress)))
            achieved = cal.end_date(task.start, done)
        points.append((task, achieved))
    return points
