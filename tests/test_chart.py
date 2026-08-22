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
    # 201 は予定終了 5/22 を過ぎて未完了なので、遅れとして数える
    assert totals["delayed"] == 1
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


def test_an_unfinished_actual_row_draws_no_actual_bar(model):
    """実績終了日が空の行は、実績のバーを描かない。

    終わっていない作業に残った実績日数は元データの書き間違いなので、
    そこから終了日を作らない (実際より進んで見えてしまうため)。
    """
    row = _by_no(model, "201")
    assert row["actual_end"] is None
    assert row["actual_days"] is None            # 書かれていた日数は捨てる
    assert row["actual"] is None                 # バーは描かない
    assert row["actual_start"] == "2026-05-13"   # 着手済みであることは残る
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


#: 状態ごとの配色 (添付ファイルから採ったもの)
STATUS_COLOURS = {
    "done": "#C0C0C0", "delayed": "#FF99CC",
    "remaining": "#FFCC00", "upcoming": "#CCFFFF",
}


def _status_row(make_filled, name, **cells):
    """1 行だけの WBS を作って、その行の描画モデルを返す。"""
    defaults = dict(start=dt.date(2026, 6, 1), days=10, end=dt.date(2026, 6, 12),
                    actual_start=None, actual_days=None, actual_end=None,
                    progress=None, member="設計")
    defaults.update(cells)
    path = make_filled(f"{name}.xlsx", rows=[(
        "開発", "", "1", name, defaults["start"], defaults["days"], defaults["end"],
        defaults["actual_start"], defaults["actual_days"], defaults["actual_end"],
        defaults["progress"], None, "", defaults["member"], "",
    )])
    return build_chart(read(path), base_date=BASE)["rows"][0]


@pytest.mark.parametrize("name,cells,expected,kind", [
    # 実績終了日が入っていれば完了
    ("完了", {"actual_start": dt.date(2026, 6, 1),
              "actual_end": dt.date(2026, 6, 8)}, "完了", "done"),
    # 着手済みで、予定終了日まであと何日か
    ("残り", {"actual_start": dt.date(2026, 6, 1)}, "残り 2 日", "remaining"),
    # 未着手で、予定開始日まであと何日か
    ("あと", {"start": dt.date(2026, 6, 22), "days": 5,
              "end": dt.date(2026, 6, 26)}, "あと 8 日", "upcoming"),
    # 予定終了日を過ぎて未完了
    ("遅れ", {"start": dt.date(2026, 5, 1), "days": 5, "end": dt.date(2026, 5, 7),
              "actual_start": dt.date(2026, 5, 1)}, "遅れ 24 日", "delayed"),
])
def test_status_is_calculated(make_filled, name, cells, expected, kind):
    """状態は「完了 / 遅れ n 日 / 残り n 日 / あと n 日」を基準日から求める。"""
    row = _status_row(make_filled, name, **cells)
    assert row["status"] == expected
    assert row["status_bg"] == STATUS_COLOURS[kind]


def test_a_row_without_dates_keeps_what_was_written(make_filled):
    """予定の日付が無ければ判断できないので、書かれた状態をそのまま残す。"""
    path = make_filled("kept.xlsx", rows=[
        ("開発", "", "1", "保留中", None, None, None,
         None, None, None, None, None, "", "設計", "保留"),
    ])
    row = build_chart(read(path), base_date=BASE)["rows"][0]
    assert row["status"] == "保留"
    assert row["status_bg"] is None      # 当てはまる色は無い


def test_the_delay_column_matches_the_status(make_filled):
    row = _status_row(make_filled, "遅れ", start=dt.date(2026, 5, 1), days=5,
                      end=dt.date(2026, 5, 7), actual_start=dt.date(2026, 5, 1))
    assert row["delay"] == 24
    assert row["status"] == "遅れ 24 日"


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
