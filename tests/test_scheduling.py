import datetime as dt

import pytest

from wbsgen.loader import from_dict
from wbsgen.scheduling import resolve_project


def build(tasks, **chart):
    chart.setdefault("base_date", dt.date(2026, 8, 7))
    chart.setdefault("start", dt.date(2026, 2, 1))
    project = from_dict({
        "project": {"title": "t"},
        "chart": chart,
        "members": [{"name": "DS"}],
        "tasks": tasks,
    })
    resolve_project(project)
    return project


def test_end_date_and_status_for_completed_task():
    task = build([{
        "name": "開発環境", "start": "2026-04-13", "days": 31,
        "actual_start": "2026-04-13", "actual_days": 75, "progress": 1.0,
    }]).tasks[0]
    assert task.end == dt.date(2026, 5, 29)
    assert task.actual_end == dt.date(2026, 7, 31)
    assert task.status == "完了"
    assert task.delay is None


def test_delay_matches_reference_workbook():
    # 添付 WBS の 38 行目: 予定終了 6/16、現在日 8/7 で「遅れ 37 日」
    task = build([{
        "name": "D-Confia_SDK機能_Andoroid", "start": "2026-04-20", "days": 38,
        "actual_start": "2026-05-06", "progress": 0.95,
    }]).tasks[0]
    assert task.end == dt.date(2026, 6, 16)
    assert task.delay == 37
    assert task.status == "遅れ 37 日"


def test_in_progress_actual_days_count_up_to_base_date():
    task = build([{
        "name": "実行中", "start": "2026-07-13", "days": 34,
        "actual_start": "2026-07-13", "progress": 0.1,
    }]).tasks[0]
    assert task.actual_days == 19
    assert task.actual_end is None
    assert task.status == "実行中"


def test_status_remaining_when_end_is_near():
    task = build([{
        "name": "残り", "start": "2026-08-03", "days": 5,
        "actual_start": "2026-08-03", "progress": 0.8,
    }]).tasks[0]
    assert task.status == "残り 0 日"


def test_status_near_when_not_started_yet():
    task = build([{"name": "あと", "start": "2026-08-14", "days": 15}]).tasks[0]
    assert task.status == "あと 4 日"


def test_status_none_when_far_in_the_future():
    task = build([{"name": "未着手", "start": "2026-12-01", "days": 10}]).tasks[0]
    assert task.status == "-"


def test_start_delay_when_task_never_started():
    task = build([{"name": "未着手遅れ", "start": "2026-08-03", "days": 20}]).tasks[0]
    assert task.delay == 4
    assert task.status == "遅れ 4 日"


def test_autoschedule_follows_predecessor():
    project = build([
        {"no": "01", "name": "A", "start": "2026-04-01", "days": 3},
        {"no": "02", "name": "B", "predecessor": "01", "days": 3},
        {"no": "03", "name": "C", "predecessor": "02", "days": 1},
    ])
    a, b, c = project.tasks
    assert (a.start, a.end) == (dt.date(2026, 4, 1), dt.date(2026, 4, 3))
    assert (b.start, b.end) == (dt.date(2026, 4, 6), dt.date(2026, 4, 8))
    assert (c.start, c.end) == (dt.date(2026, 4, 9), dt.date(2026, 4, 9))


def test_autoschedule_uses_latest_of_multiple_predecessors():
    project = build([
        {"no": "01", "name": "A", "start": "2026-04-01", "days": 3},
        {"no": "02", "name": "B", "start": "2026-04-01", "days": 8},
        {"no": "03", "name": "C", "predecessor": "01,02", "days": 1},
    ])
    assert project.tasks[2].start == dt.date(2026, 4, 13)


def test_circular_predecessor_is_rejected():
    with pytest.raises(ValueError, match="循環"):
        build([
            {"no": "01", "name": "A", "predecessor": "02", "days": 1},
            {"no": "02", "name": "B", "predecessor": "01", "days": 1},
        ])


def test_summary_rolls_up_children():
    project = build([
        {"group": "設計", "no": "10", "name": "設計", "kind": "summary"},
        {"no": "11", "name": "画面", "start": "2026-04-01", "days": 5, "effort": 5,
         "progress": 1.0},
        {"no": "12", "name": "I/F", "start": "2026-04-06", "days": 5, "effort": 3,
         "progress": 0.0},
    ])
    summary = project.tasks[0]
    assert summary.start == dt.date(2026, 4, 1)
    assert summary.end == dt.date(2026, 4, 10)
    assert summary.effort == 8
    assert summary.progress == 0.5
    assert summary.status == "-"
