"""仕様 (BlankWBS) の組み立てのテスト。"""

import datetime as dt

import pytest

from wbsgen.blank import (
    DEFAULT_ROWS,
    SpecError,
    add_months,
    build,
    from_dict,
    parse_date,
    to_dict,
)


# ---------------------------------------------------------------- 期間
def test_period_from_end_date():
    spec = build(dt.date(2026, 4, 1), end=dt.date(2027, 3, 31))
    assert spec.period_days == 365
    assert spec.end == dt.date(2027, 3, 31)


def test_period_from_months():
    spec = build(dt.date(2026, 4, 1), months=6)
    assert spec.end == dt.date(2026, 9, 30)


def test_period_from_days():
    spec = build(dt.date(2026, 4, 1), period_days=10)
    assert spec.end == dt.date(2026, 4, 10)


def test_period_defaults_to_a_year():
    spec = build(dt.date(2026, 4, 1))
    assert spec.period_days == 365
    assert spec.rows == DEFAULT_ROWS
    assert spec.unit == "week"


@pytest.mark.parametrize("start,months,expected", [
    ((2026, 1, 31), 1, (2026, 2, 28)),
    ((2028, 1, 31), 1, (2028, 2, 29)),
    ((2026, 12, 1), 3, (2027, 3, 1)),
])
def test_add_months_clamps_to_the_month_end(start, months, expected):
    assert add_months(dt.date(*start), months) == dt.date(*expected)


# ---------------------------------------------------------------- 検証
@pytest.mark.parametrize("kwargs,message", [
    ({"end": dt.date(2026, 3, 1)}, "終了日"),
    ({"period_days": 0}, "期間"),
    ({"period_days": 99999}, "期間"),
    ({"rows": -1}, "行数"),
    ({"rows": 99999}, "行数"),
    ({"unit": "hour"}, "表示単位"),
    ({"workdays": ["someday"]}, "曜日"),
    ({"members": [f"m{i}" for i in range(200)]}, "担当者"),
])
def test_invalid_input_is_rejected(kwargs, message):
    with pytest.raises(SpecError, match=message):
        build(dt.date(2026, 4, 1), **kwargs)


def test_title_defaults_to_the_start_year():
    assert build(dt.date(2026, 4, 1)).title == "2026年 スケジュール"
    assert build(dt.date(2026, 4, 1), title="  ").title == "2026年 スケジュール"
    assert build(dt.date(2026, 4, 1), title=" 計画 ").title == "計画"


def test_blank_member_names_are_dropped():
    spec = build(dt.date(2026, 4, 1), members=["設計", "  ", "", "製造"])
    assert spec.members == ["設計", "製造"]


# ---------------------------------------------------------------- カレンダー
def test_calendar_includes_japanese_holidays_by_default():
    cal = build(dt.date(2026, 4, 1), months=2).calendar()
    assert not cal.is_workday(dt.date(2026, 4, 29))
    assert cal.is_workday(dt.date(2026, 4, 30))


def test_japanese_holidays_can_be_turned_off():
    cal = build(dt.date(2026, 4, 1), months=2, japanese_holidays=False).calendar()
    assert cal.is_workday(dt.date(2026, 4, 29))


def test_extra_holidays_are_applied():
    cal = build(dt.date(2026, 4, 1), months=2,
                holidays=[dt.date(2026, 4, 30)]).calendar()
    assert not cal.is_workday(dt.date(2026, 4, 30))


# ---------------------------------------------------------------- 辞書
def test_round_trip_through_dict():
    spec = build(dt.date(2026, 4, 1), end=dt.date(2027, 3, 31), unit="day",
                 rows=25, title="計画", members=["設計"],
                 workdays=["mon", "sat"], japanese_holidays=False,
                 holidays=[dt.date(2026, 12, 30)])
    assert to_dict(from_dict(to_dict(spec))) == to_dict(spec)


def test_from_dict_accepts_months_and_slashes():
    spec = from_dict({"start": "2026/04/01", "months": 3})
    assert spec.start == dt.date(2026, 4, 1)
    assert spec.end == dt.date(2026, 6, 30)


@pytest.mark.parametrize("payload,message", [
    ({}, "start"),
    ({"start": "bad"}, "start"),
    ({"start": "2026-04-01", "rows": "many"}, "rows"),
    ({"start": "2026-04-01", "holidays": ["nope"]}, "holidays"),
])
def test_from_dict_rejects_bad_payloads(payload, message):
    with pytest.raises(SpecError, match=message):
        from_dict(payload)


def test_from_dict_rejects_a_non_mapping():
    with pytest.raises(SpecError, match="マッピング"):
        from_dict([1, 2, 3])


def test_parse_date_accepts_date_objects():
    assert parse_date(dt.date(2026, 4, 1), "start") == dt.date(2026, 4, 1)
    assert parse_date(dt.datetime(2026, 4, 1, 9, 0), "start") == dt.date(2026, 4, 1)
