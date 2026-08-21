import datetime as dt

import pytest

from wbsgen.timeline import UNIT_DAY, UNIT_MONTH, UNIT_WEEK, Timeline
from wbsgen.workcal import WorkCalendar


@pytest.fixture
def cal():
    return WorkCalendar.build()


def test_week_columns_step_from_the_start_date(cal):
    """元ファイルと同じく、週は曜日ではなく表示開始日を起点に 7 日刻みで並ぶ。"""
    tl = Timeline(dt.date(2026, 2, 1), 28, UNIT_WEEK, cal)   # 2026-02-01 は日曜
    assert [c.start for c in tl.columns] == [
        dt.date(2026, 2, 1), dt.date(2026, 2, 8),
        dt.date(2026, 2, 15), dt.date(2026, 2, 22),
    ]
    assert tl.start == dt.date(2026, 2, 1)
    assert tl.end == dt.date(2026, 2, 28)


def test_day_columns_are_one_day_each(cal):
    tl = Timeline(dt.date(2026, 4, 1), 5, UNIT_DAY, cal)
    assert [c.start for c in tl.columns] == [
        dt.date(2026, 4, d) for d in range(1, 6)]
    assert all(c.days == 1 for c in tl.columns)


def test_month_columns_snap_to_calendar_months(cal):
    tl = Timeline(dt.date(2026, 4, 15), 90, UNIT_MONTH, cal)
    assert tl.start == dt.date(2026, 4, 1)
    assert [c.start.month for c in tl.columns] == [4, 5, 6, 7]
    assert tl.columns[0].end == dt.date(2026, 4, 30)


# ----------------------------------------------------------------------
def test_week_header_matches_the_reference_workbook(cal):
    """上段は月、下段は週の開始日。上段は区切りが変わる列にだけ値が入る。"""
    tl = Timeline(dt.date(2026, 2, 1), 150, UNIT_WEEK, cal)
    assert tl.formats == ('m"月"', "m/d")

    # 添付ファイルの S3 / W3 / AB3 / AF3 と同じ位置・同じ日付
    top = [(index, day, text) for index, _span, day, text in tl.header_top()]
    assert top[:4] == [
        (0, dt.date(2026, 2, 1), "2月"),
        (4, dt.date(2026, 3, 1), "3月"),
        (9, dt.date(2026, 4, 5), "4月"),
        (13, dt.date(2026, 5, 3), "5月"),
    ]

    bottom = tl.header_bottom()
    assert len(bottom) == len(tl)
    assert bottom[:3] == [
        (0, dt.date(2026, 2, 1), "2/1"),
        (1, dt.date(2026, 2, 8), "2/8"),
        (2, dt.date(2026, 2, 15), "2/15"),
    ]


def test_day_header_has_month_and_weekday(cal):
    tl = Timeline(dt.date(2026, 4, 1), 40, UNIT_DAY, cal)
    assert tl.formats == ('m"月"', "d")
    assert [t for _, _, _, t in tl.header_top()] == ["4月", "5月"]
    assert tl.column_label(tl.columns[0]) == "1"
    assert tl.weekday_label(tl.columns[0]) == "水"


def test_month_header_has_year_and_month(cal):
    tl = Timeline(dt.date(2026, 11, 1), 120, UNIT_MONTH, cal)
    assert tl.formats == ('yyyy"年"', 'm"月"')
    assert [t for _, _, _, t in tl.header_top()] == ["2026年", "2027年"]
    assert tl.column_label(tl.columns[0]) == "11月"
    assert tl.weekday_label(tl.columns[0]) == ""


# ----------------------------------------------------------------------
def test_rest_columns_are_flagged_only_for_day_unit():
    cal = WorkCalendar.build(holidays=[dt.date(2026, 4, 29)])
    day = Timeline(dt.date(2026, 4, 1), 30, UNIT_DAY, cal)
    rest = [c.start for c in day.columns if day.is_rest_column(c)]
    assert dt.date(2026, 4, 4) in rest      # 土
    assert dt.date(2026, 4, 5) in rest      # 日
    assert dt.date(2026, 4, 29) in rest     # 休業日
    assert dt.date(2026, 4, 30) not in rest
    assert [c.start for c in day.columns if day.is_saturday_column(c)] == [
        dt.date(2026, 4, 4), dt.date(2026, 4, 11),
        dt.date(2026, 4, 18), dt.date(2026, 4, 25),
    ]

    week = Timeline(dt.date(2026, 4, 1), 30, UNIT_WEEK, cal)
    assert not any(week.is_rest_column(c) for c in week.columns)
