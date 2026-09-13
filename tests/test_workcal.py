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

# ------------------------------------------------- 日付が動いた年 (内閣府の一覧)
#: 祝日法どおりに求められない年。内閣府「国民の祝日について」の一覧そのまま。
SPECIAL = {
    # 天皇の即位に伴う祝日。この年に天皇誕生日 (12/23) は無い
    2019: ["2019-01-01", "2019-01-14", "2019-02-11", "2019-03-21", "2019-04-29",
           "2019-04-30", "2019-05-01", "2019-05-02", "2019-05-03", "2019-05-04",
           "2019-05-05", "2019-05-06", "2019-07-15", "2019-08-11", "2019-08-12",
           "2019-09-16", "2019-09-23", "2019-10-14", "2019-10-22", "2019-11-03",
           "2019-11-04", "2019-11-23"],
    # オリンピックで海の日・スポーツの日・山の日が動いた年
    2020: ["2020-01-01", "2020-01-13", "2020-02-11", "2020-02-23", "2020-02-24",
           "2020-03-20", "2020-04-29", "2020-05-03", "2020-05-04", "2020-05-05",
           "2020-05-06", "2020-07-23", "2020-07-24", "2020-08-10", "2020-09-21",
           "2020-09-22", "2020-11-03", "2020-11-23"],
    2021: ["2021-01-01", "2021-01-11", "2021-02-11", "2021-02-23", "2021-03-20",
           "2021-04-29", "2021-05-03", "2021-05-04", "2021-05-05", "2021-07-22",
           "2021-07-23", "2021-08-08", "2021-08-09", "2021-09-20", "2021-09-23",
           "2021-11-03", "2021-11-23"],
}


@pytest.mark.parametrize("year", sorted(SPECIAL))
def test_the_years_whose_dates_moved_match_the_official_list(year):
    assert [d.isoformat() for d in japanese_holidays(year)] == SPECIAL[year]


@pytest.mark.parametrize("year, expected", [
    (2024, ["2024-01-01", "2024-01-08", "2024-02-11", "2024-02-12", "2024-02-23",
            "2024-03-20", "2024-04-29", "2024-05-03", "2024-05-04", "2024-05-05",
            "2024-05-06", "2024-07-15", "2024-08-11", "2024-08-12", "2024-09-16",
            "2024-09-22", "2024-09-23", "2024-10-14", "2024-11-03", "2024-11-04",
            "2024-11-23"]),
    (2025, ["2025-01-01", "2025-01-13", "2025-02-11", "2025-02-23", "2025-02-24",
            "2025-03-20", "2025-04-29", "2025-05-03", "2025-05-04", "2025-05-05",
            "2025-05-06", "2025-07-21", "2025-08-11", "2025-09-15", "2025-09-23",
            "2025-10-13", "2025-11-03", "2025-11-23", "2025-11-24"]),
    (2027, ["2027-01-01", "2027-01-11", "2027-02-11", "2027-02-23", "2027-03-21",
            "2027-03-22", "2027-04-29", "2027-05-03", "2027-05-04", "2027-05-05",
            "2027-07-19", "2027-08-11", "2027-09-20", "2027-09-23", "2027-10-11",
            "2027-11-03", "2027-11-23"]),
])
def test_ordinary_years_match_the_official_list(year, expected):
    """祝日法どおりの年。春分・秋分、ハッピーマンデー、振替まで合うこと。"""
    assert [d.isoformat() for d in japanese_holidays(year)] == expected


def test_new_years_day_on_a_sunday_moves_to_the_second():
    """元日が日曜なら 1/2 が振替休日になる。"""
    days = japanese_holidays(2023)
    assert dt.date(2023, 1, 1) in days and dt.date(2023, 1, 2) in days
