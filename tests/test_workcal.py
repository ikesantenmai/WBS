import datetime as dt

import pytest

from wbsgen.workcal import (
    WorkCalendar,
    japanese_holidays,
    japanese_holidays_range,
    parse_weekdays,
)


@pytest.fixture
def cal():
    holidays = japanese_holidays_range(dt.date(2026, 1, 1), dt.date(2027, 12, 31))
    return WorkCalendar.build(holidays=holidays)


def test_parse_weekdays_accepts_english_and_japanese():
    assert parse_weekdays(["mon", "fri"]) == frozenset({0, 4})
    assert parse_weekdays(["土", "日"]) == frozenset({5, 6})
    with pytest.raises(ValueError):
        parse_weekdays(["someday"])


def test_weekend_and_holiday_are_not_workdays(cal):
    assert cal.is_workday(dt.date(2026, 4, 13))       # 月
    assert not cal.is_workday(dt.date(2026, 4, 18))   # 土
    assert not cal.is_workday(dt.date(2026, 4, 29))   # 昭和の日


def test_extra_workday_overrides_holiday():
    cal = WorkCalendar.build(holidays=[dt.date(2026, 5, 3)],
                             extra_workdays=[dt.date(2026, 5, 3)])
    assert cal.is_workday(dt.date(2026, 5, 3))


# 添付の WBS から実測した (開始日, 稼働日数, 終了日) の組
@pytest.mark.parametrize("start,days,end", [
    ((2026, 4, 13), 31, (2026, 5, 29)),
    ((2026, 10, 1), 40, (2026, 11, 30)),
    ((2026, 4, 3), 16, (2026, 4, 24)),
    ((2026, 4, 3), 27, (2026, 5, 15)),
    ((2026, 4, 2), 20, (2026, 4, 30)),
    ((2026, 4, 20), 38, (2026, 6, 16)),
    ((2026, 4, 1), 39, (2026, 5, 29)),
    ((2026, 8, 3), 5, (2026, 8, 7)),
])
def test_end_date_matches_reference_workbook(cal, start, days, end):
    assert cal.end_date(dt.date(*start), days) == dt.date(*end)


def test_end_date_starts_from_next_workday(cal):
    # 2027-01-11 は成人の日。翌稼働日から数え始める。
    assert cal.end_date(dt.date(2027, 1, 11), 9) == dt.date(2027, 1, 22)


def test_workdays_between_counts_both_ends(cal):
    assert cal.workdays_between(dt.date(2026, 8, 3), dt.date(2026, 8, 7)) == 5
    assert cal.workdays_between(dt.date(2026, 6, 16), dt.date(2026, 8, 7)) == 38
    assert cal.workdays_between(dt.date(2026, 8, 7), dt.date(2026, 8, 3)) == 0


def test_japanese_holidays_include_substitutes():
    days = japanese_holidays(2026)
    assert dt.date(2026, 5, 6) in days     # 5/3 が日曜のため振替
    assert dt.date(2026, 9, 22) in days    # 国民の休日
    assert dt.date(2026, 2, 23) in days    # 天皇誕生日 (2020 年以降)
    assert dt.date(2026, 12, 23) not in days
