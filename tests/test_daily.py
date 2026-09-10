"""本日の状況 (日次チェック) のテスト。

読み込んだ WBS を基準日から見て、その日に手を打つべき行を拾えること。
"""

import datetime as dt

import pytest

from wbsgen.daily import build
from wbsgen.importer import Row, resolve
from wbsgen.workcal import WorkCalendar

D = dt.date
TODAY = D(2026, 9, 10)          # 木曜日


def _calendar():
    return WorkCalendar()


def _row(number, name="作業", **kwargs):
    row = Row(row=number, name=name, **kwargs)
    row.written = {key: getattr(row, key) for key in
                   ("days", "end", "actual_days", "actual_end", "progress",
                    "delay", "status")}
    return row


def _build(rows, base=TODAY):
    """基準日で解決してから数える (画面と同じ順序)。"""
    calendar = _calendar()
    resolve(rows, calendar, base)
    return build(rows, calendar, base)


# ---------------------------------------------------------------- 本日の予定
def test_tasks_planned_to_start_today_are_listed():
    rows = [
        _row(2, start=TODAY, end=D(2026, 9, 18)),
        _row(3, start=D(2026, 9, 9), end=D(2026, 9, 18)),
        _row(4, start=TODAY, end=TODAY),
    ]
    assert _build(rows).starting == [2, 4]


def test_tasks_planned_to_end_today_are_listed():
    rows = [
        _row(2, start=D(2026, 9, 1), end=TODAY),
        _row(3, start=D(2026, 9, 1), end=D(2026, 9, 11)),
    ]
    assert _build(rows).ending == [2]


def test_a_summary_row_counts_too():
    """大項目だけの行も数える (画面の集計と件数をそろえるため)。"""
    rows = [_row(2, name="", group="ITa", start=TODAY, end=TODAY)]
    assert _build(rows).starting == [2]


def test_a_row_with_only_a_number_still_counts():
    rows = [_row(2, name="", no="1", start=TODAY, end=TODAY)]
    assert _build(rows).starting == [2]


# ---------------------------------------------------------------- 遅延
def test_delayed_tasks_are_listed_worst_first():
    rows = [
        _row(2, start=D(2026, 9, 8), end=D(2026, 9, 30)),      # 開始遅れ 2 日
        _row(3, start=D(2026, 8, 3), end=D(2026, 8, 31)),      # もっと遅れ
        _row(4, start=TODAY, end=D(2026, 9, 30)),              # 遅れなし
    ]
    digest = _build(rows)
    assert digest.delayed == [3, 2]


def test_a_finished_task_is_not_delayed():
    rows = [_row(2, start=D(2026, 8, 3), end=D(2026, 8, 31),
                 actual_start=D(2026, 8, 3), actual_end=D(2026, 9, 4))]
    assert _build(rows).delayed == []


# ---------------------------------------------------------------- 未着手
def test_tasks_without_any_actual_are_not_started():
    rows = [
        _row(2, start=D(2026, 9, 20), end=D(2026, 9, 25)),
        _row(3, start=D(2026, 9, 1), end=D(2026, 9, 30),
             actual_start=D(2026, 9, 1)),
        _row(4, start=D(2026, 9, 14), end=D(2026, 9, 18)),
    ]
    # 予定の開始日が早い順
    assert _build(rows).not_started == [4, 2]


def test_a_row_with_no_plan_is_not_counted_as_not_started():
    """日付が 1 つも無い行は、着手の話ではなく書き漏れとして拾う。"""
    rows = [_row(2)]
    digest = _build(rows)
    assert digest.not_started == []
    assert _found(digest, "no_dates") == [2]


def test_a_finished_task_is_not_counted_as_not_started():
    """実績の終了日だけ書いてあることがある (開始日を書き忘れた行)。"""
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 3),
                 actual_end=D(2026, 9, 3))]
    assert _build(rows).not_started == []


# ---------------------------------------------------------------- 整合性
def _found(digest, kind):
    for check in digest.checks:
        if check.kind == kind:
            return check.rows
    raise AssertionError(f"{kind} を確かめていない")


def test_every_check_is_reported_even_when_nothing_is_wrong():
    """0 件のチェックも返す (何を確かめたかが画面で判るように)。"""
    digest = _build([_row(2, start=D(2026, 9, 1), end=D(2026, 9, 3),
                          actual_start=D(2026, 9, 1), actual_end=D(2026, 9, 3))])
    assert len(digest.checks) >= 10
    assert all(check.rows == [] for check in digest.checks)


def test_a_plan_that_ends_before_it_starts_is_found():
    rows = [_row(2, start=D(2026, 9, 10), end=D(2026, 9, 3))]
    assert _found(_build(rows), "end_before_start") == [2]


def test_an_actual_that_ends_before_it_starts_is_found():
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 30),
                 actual_start=D(2026, 9, 8), actual_end=D(2026, 9, 3))]
    assert _found(_build(rows), "actual_end_before_start") == [2]


def test_a_plan_with_only_one_date_is_found():
    """終了日が無く、日数からも決まらない行。"""
    rows = [_row(2, end=D(2026, 9, 30))]
    assert _found(_build(rows), "plan_incomplete") == [2]


def test_a_plan_completed_by_the_number_of_days_is_not_flagged():
    rows = [_row(2, start=D(2026, 9, 1), days=5)]
    assert _found(_build(rows), "plan_incomplete") == []


def test_an_actual_end_without_a_start_is_found():
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 3),
                 actual_end=D(2026, 9, 3))]
    assert _found(_build(rows), "actual_end_without_start") == [2]


def test_progress_without_an_actual_start_is_found():
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 30), progress=0.5)]
    assert _found(_build(rows), "progress_without_actual") == [2]


def test_full_progress_without_an_actual_end_is_found():
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 30),
                 actual_start=D(2026, 9, 1), progress=1.0)]
    assert _found(_build(rows), "done_without_actual_end") == [2]


def test_a_written_number_of_days_that_does_not_match_is_found():
    """開始日・終了日から数え直した日数と、書かれた日数が違う行。"""
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 4), days=99)]
    assert _found(_build(rows), "days_rewritten") == [2]


def test_a_matching_number_of_days_is_not_flagged():
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 4), days=4)]
    assert _found(_build(rows), "days_rewritten") == []


def test_an_actual_number_of_days_left_behind_is_found():
    """実績の終了日が無いのに日数だけ残っている (元データの書き間違い)。"""
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 30),
                 actual_start=D(2026, 9, 1), actual_days=3)]
    assert _found(_build(rows), "actual_days_rewritten") == [2]


def test_an_unknown_predecessor_is_found():
    rows = [
        _row(2, no="1", start=D(2026, 9, 1), end=D(2026, 9, 4)),
        _row(3, no="2", start=D(2026, 9, 7), end=D(2026, 9, 9), predecessor="9"),
    ]
    assert _found(_build(rows), "predecessor_missing") == [3]


def test_a_task_starting_before_its_predecessor_ends_is_found():
    rows = [
        _row(2, no="1", start=D(2026, 9, 1), end=D(2026, 9, 18)),
        _row(3, no="2", start=D(2026, 9, 7), end=D(2026, 9, 9), predecessor="1"),
        _row(4, no="3", start=D(2026, 9, 21), end=D(2026, 9, 24), predecessor="1"),
    ]
    digest = _build(rows)
    assert _found(digest, "predecessor_order") == [3]
    assert _found(digest, "predecessor_missing") == []


@pytest.mark.parametrize("text", ["1, 2", "1・2", "1 2", "1、2"])
def test_predecessors_can_be_written_in_several_ways(text):
    rows = [
        _row(2, no="1", start=D(2026, 9, 1), end=D(2026, 9, 4)),
        _row(3, no="2", start=D(2026, 9, 1), end=D(2026, 9, 4)),
        _row(4, no="3", start=D(2026, 9, 7), end=D(2026, 9, 9), predecessor=text),
    ]
    assert _found(_build(rows), "predecessor_missing") == []


# ---------------------------------------------------------------- 担当
def test_owners_without_a_task_today_are_listed():
    rows = [
        _row(2, start=D(2026, 9, 1), end=D(2026, 9, 30), member="高瀬"),
        _row(3, start=D(2026, 8, 1), end=D(2026, 8, 31), member="吉田"),
        _row(4, start=D(2026, 9, 20), end=D(2026, 9, 25), member="鈴木さん"),
    ]
    digest = _build(rows)
    assert digest.members == ["高瀬", "吉田", "鈴木"]
    assert digest.idle_members == ["吉田", "鈴木"]


def test_an_owner_named_in_a_shared_cell_counts_as_assigned():
    """「佐々木（高瀬）」のように 1 つのセルに複数書いてあっても数える。"""
    rows = [
        _row(2, start=D(2026, 9, 1), end=D(2026, 9, 30), member="佐々木（高瀬）"),
        _row(3, start=D(2026, 8, 1), end=D(2026, 8, 31), member="高瀬"),
    ]
    assert _build(rows).idle_members == []


def test_nobody_is_idle_when_no_owner_is_written():
    rows = [_row(2, start=D(2026, 9, 1), end=D(2026, 9, 30))]
    digest = _build(rows)
    assert digest.members == []
    assert digest.idle_members == []


# ---------------------------------------------------------------- 基準日
def test_the_base_date_is_reported():
    assert _build([_row(2, start=TODAY, end=TODAY)]).date == TODAY


def test_another_base_date_changes_what_is_due():
    rows = [_row(2, start=D(2026, 9, 14), end=D(2026, 9, 18))]
    assert _build(rows, base=D(2026, 9, 14)).starting == [2]
    assert _build(rows, base=D(2026, 9, 14)).delayed == []
