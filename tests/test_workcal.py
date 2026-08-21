import datetime as dt

import pytest

from wbsgen.workcal import (
    WorkCalendar,
    japanese_holidays,
    japanese_holidays_range,
    parse_weekdays,
)


def test_parse_weekdays_accepts_english_and_japanese():
    assert parse_weekdays(["mon", "fri"]) == frozenset({0, 4})
    assert parse_weekdays(["土", "日"]) == frozenset({5, 6})
    with pytest.raises(ValueError, match="曜日"):
        parse_weekdays(["someday"])


def test_weekend_and_holiday_are_not_workdays():
    cal = WorkCalendar.build(holidays=japanese_holidays(2026))
    assert cal.is_workday(dt.date(2026, 4, 13))       # 月
    assert not cal.is_workday(dt.date(2026, 4, 18))   # 土
    assert not cal.is_workday(dt.date(2026, 4, 29))   # 昭和の日


def test_workdays_can_include_saturday():
    cal = WorkCalendar.build(workdays=["mon", "tue", "wed", "thu", "fri", "sat"])
    assert cal.is_workday(dt.date(2026, 4, 18))
    assert not cal.is_workday(dt.date(2026, 4, 19))


def test_extra_workday_overrides_holiday():
    cal = WorkCalendar.build(holidays=[dt.date(2026, 5, 3)],
                             extra_workdays=[dt.date(2026, 5, 3)])
    assert cal.is_workday(dt.date(2026, 5, 3))


def test_holidays_between_lists_only_registered_holidays():
    cal = WorkCalendar.build(holidays=[dt.date(2026, 5, 3), dt.date(2027, 1, 1)])
    assert cal.holidays_between(dt.date(2026, 1, 1), dt.date(2026, 12, 31)) == [
        dt.date(2026, 5, 3)]


def test_japanese_holidays_include_substitutes():
    days = japanese_holidays(2026)
    assert dt.date(2026, 5, 6) in days     # 5/3 が日曜のため振替
    assert dt.date(2026, 9, 22) in days    # 国民の休日
    assert dt.date(2026, 2, 23) in days    # 天皇誕生日 (2020 年以降)
    assert dt.date(2026, 12, 23) not in days


def test_japanese_holidays_range_matches_the_reference_workbook():
    """添付ファイルの休日設定 (2026-02-01 以降) と一致すること。"""
    days = japanese_holidays_range(dt.date(2026, 2, 1), dt.date(2026, 12, 31))
    assert [d.isoformat() for d in days] == [
        "2026-02-11", "2026-02-23", "2026-03-20", "2026-04-29", "2026-05-03",
        "2026-05-04", "2026-05-05", "2026-05-06", "2026-07-20", "2026-08-11",
        "2026-09-21", "2026-09-22", "2026-09-23", "2026-10-12", "2026-11-03",
        "2026-11-23",
    ]
