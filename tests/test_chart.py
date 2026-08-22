"""ガントチャートの描画モデルのテスト。"""

import datetime as dt

import pytest

from wbsgen.chart import build as build_chart
from wbsgen.importer import read

BASE = dt.date(2026, 6, 10)


@pytest.fixture
def model(filled_book):
    return build_chart(read(filled_book, "filled.xlsx"), base_date=BASE)


def _by_no(model, no):
    return next(r for r in model["rows"] if r["no"] == no)


# ---------------------------------------------------------------- 全体
def test_timeline_comes_from_the_file(model):
    timeline = model["timeline"]
    assert timeline["unit"] == "week"
    assert timeline["start"] == "2026-04-01"
    assert [c["label"] for c in timeline["columns"][:3]] == ["4/1", "4/8", "4/15"]
    assert [b["label"] for b in timeline["bands"]][:3] == ["4月", "5月", "6月"]


def test_now_line_is_placed_on_the_base_date(model):
    # 2026-06-10 は表示開始 4/1 から 10 週目
    assert model["now_x"] == pytest.approx(10.0)


def test_now_line_is_absent_outside_the_period(filled_book):
    model = build_chart(read(filled_book), base_date=dt.date(2030, 1, 1))
    assert model["now_x"] is None


def test_totals(model):
    totals = model["totals"]
    assert totals["rows"] == 5
    assert totals["done"] == 2          # 実績終了日が入っている 2 行
    assert totals["running"] == 2
    assert totals["delayed"] == 0
    assert totals["effort"] == 105.0
    assert totals["first_day"] == "2026-04-01"
    # 期間の終わりは、終了日が空の行を補った値まで含む (バーと一致させる)
    assert totals["last_day"] == "2026-08-17"


def test_overall_progress_is_weighted_by_the_counted_days(model):
    """日数は日付から数え直されるので、重みもその値になる。"""
    rows = model["rows"]
    weight = sum(r["days"] for r in rows)
    expected = sum((r["progress"] or 0) * r["days"] for r in rows) / weight
    assert model["totals"]["progress"] == pytest.approx(expected, abs=1e-4)


# ---------------------------------------------------------------- バー
def test_plan_and_actual_bars(model):
    row = _by_no(model, "101")
    # 4/1〜4/14 は 1 列目の頭から 3 列目の頭まで (週表示)
    assert row["plan"] == {"x1": 0.0, "x2": pytest.approx(2.0)}
    assert row["actual"] == {"x1": 0.0, "x2": pytest.approx(2.0)}


def test_a_row_without_an_actual_start_has_no_actual_bar(model):
    assert _by_no(model, "301")["actual"] is None
    assert _by_no(model, "301")["plan"] is not None


def test_an_unfinished_actual_bar_uses_the_recorded_days(model):
    """実績終了日が空でも、実績日数からバーの右端を決める。"""
    row = _by_no(model, "201")
    assert "actual_end" in row["derived"]        # 補った値であることが判る
    assert row["actual"]["x2"] > row["actual"]["x1"]
    assert row["progress"] == 0.6                # 完了扱いにはしない


def test_a_missing_end_date_is_derived_from_the_workdays(model):
    """終了日が書かれていない行は、日数 (稼働日) から補ってバーを描く。"""
    row = _by_no(model, "301")
    assert row["end"] == "2026-08-17"       # 7/27 から 15 稼働日
    assert "end" in row["derived"]
    assert _by_no(model, "101")["derived"] == []


def test_days_are_counted_from_the_dates(model):
    """予定・実績の日数は、書かれた値ではなく日付から数える。"""
    row = _by_no(model, "102")
    assert (row["start"], row["end"]) == ("2026-04-15", "2026-05-07")
    assert row["days"] == 13                    # ファイルには 15 と書いてある
    assert (row["actual_start"], row["actual_end"]) == ("2026-04-15", "2026-05-12")
    assert row["actual_days"] == 16             # ファイルには 18 と書いてある
    assert set(row["derived"]) == {"days", "actual_days"}


def test_an_actual_end_date_makes_it_complete(model):
    assert _by_no(model, "101")["progress"] == 1.0
    assert _by_no(model, "102")["progress"] == 1.0


def test_bars_outside_the_period_are_dropped(make_filled):
    path = make_filled("out.xlsx", rows=[
        ("開発", "", "1", "範囲内", dt.date(2026, 4, 6), 5, dt.date(2026, 4, 10),
         None, None, None, None, None, "", "設計", ""),
        ("開発", "", "2", "範囲外", dt.date(2030, 1, 1), 5, dt.date(2030, 1, 7),
         None, None, None, None, None, "", "設計", ""),
    ])
    model = build_chart(read(path), base_date=BASE)
    assert _by_no(model, "1")["plan"] is not None
    assert _by_no(model, "2")["plan"] is None


# ---------------------------------------------------------------- 色
def test_members_get_stable_colours(model):
    colours = {m["name"]: m["color"] for m in model["members"]}
    assert colours["設計"] != colours["製造"] != colours["テスト"]
    assert _by_no(model, "101")["color"] == colours["設計"]
    assert _by_no(model, "201")["color"] == colours["製造"]


def test_a_member_not_on_the_list_still_gets_a_colour(make_filled):
    path = make_filled("x.xlsx", members=[], rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 6), 5, dt.date(2026, 4, 10),
         None, None, None, None, None, "", "外注", ""),
    ])
    model = build_chart(read(path), base_date=BASE)
    assert model["rows"][0]["color"].startswith("#")


@pytest.mark.parametrize("status,background", [
    ("完了", "#C0C0C0"),
    ("実行中", "#FFFF99"),
    ("遅れ 3 日", "#FF99CC"),
    ("あと 2 日", "#CCFFFF"),
    ("残り 1 日", "#FFCC00"),
])
def test_status_keeps_the_original_colours(make_filled, status, background):
    """状態の配色は添付ファイルから採ったものを使う。"""
    path = make_filled("s.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 6), 5, dt.date(2026, 4, 10),
         None, None, None, None, None, "", "設計", status),
    ])
    model = build_chart(read(path), base_date=BASE)
    assert model["rows"][0]["status_bg"] == background


def test_an_unknown_status_has_no_colour(make_filled):
    path = make_filled("u.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 6), 5, dt.date(2026, 4, 10),
         None, None, None, None, None, "", "設計", "保留"),
    ])
    model = build_chart(read(path), base_date=BASE)
    assert model["rows"][0]["status_bg"] is None


# ---------------------------------------------------------------- 表示単位
@pytest.mark.parametrize("unit,columns", [("day", 365), ("week", 53), ("month", 12)])
def test_the_unit_can_be_overridden(filled_book, unit, columns):
    imported = read(filled_book)
    imported.spec.unit = unit
    model = build_chart(imported, base_date=BASE)
    assert len(model["timeline"]["columns"]) == columns
    # 単位を変えてもバーの示す期間は同じ
    row = next(r for r in model["rows"] if r["no"] == "101")
    assert row["start"] == "2026-04-01" and row["end"] == "2026-04-14"
