"""記入済み Excel の読み込みのテスト。"""

import datetime as dt

import openpyxl
import pytest

from wbsgen.blank import SpecError, build
from wbsgen.importer import read
from wbsgen.workbook import SHEET_PLAN, write


# ---------------------------------------------------------------- 基本
def test_reads_the_rows(filled_book):
    imported = read(filled_book, "filled.xlsx")
    assert [r.name for r in imported.rows] == [
        "要件定義", "基本設計", "詳細設計", "コーディング", "結合テスト"]
    assert imported.warnings == []


def test_reads_the_values(filled_book):
    row = read(filled_book).rows[0]
    assert row.no == "101"
    assert row.name == "要件定義"
    assert row.start == dt.date(2026, 4, 1)
    assert row.days == 10
    assert row.end == dt.date(2026, 4, 14)
    assert row.actual_start == dt.date(2026, 4, 1)
    assert row.actual_end == dt.date(2026, 4, 14)
    assert row.progress == 1.0
    assert row.effort == 10
    assert row.member == "設計"
    assert row.status == "完了"


def test_group_and_subgroup_carry_down(filled_book):
    """大項目・中項目は変わった行にだけ書かれるので、下の行へ引き継ぐ。"""
    rows = read(filled_book).rows
    assert [(r.group, r.subgroup) for r in rows] == [
        ("開発", "要件"), ("開発", "要件"), ("開発", "製造"),
        ("開発", "製造"), ("テスト", "結合"),
    ]


def test_blank_rows_are_skipped(filled_book):
    """記入されていない空行は読み飛ばす。"""
    assert len(read(filled_book).rows) == 5      # 空行は 30 行ある


def test_title_comes_from_the_sheet(filled_book):
    assert read(filled_book).title == "2026年度 開発スケジュール"


# ---------------------------------------------------------------- 設定
def test_period_comes_from_the_config_sheet(filled_book):
    spec = read(filled_book).spec
    assert spec.start == dt.date(2026, 4, 1)
    assert spec.period_days == 365
    assert spec.unit == "week"


def test_members_come_from_the_member_sheet(filled_book):
    assert read(filled_book).spec.members == ["設計", "製造", "テスト"]


def test_holidays_come_from_the_config_sheet(filled_book):
    spec = read(filled_book).spec
    assert dt.date(2026, 4, 29) in spec.holidays      # 昭和の日
    assert not spec.calendar().is_workday(dt.date(2026, 4, 29))


def test_workdays_come_from_the_config_sheet(make_filled):
    path = make_filled("sat.xlsx", workdays=["mon", "tue", "wed", "thu", "fri", "sat"])
    spec = read(path).spec
    assert spec.workdays == ["mon", "tue", "wed", "thu", "fri", "sat"]
    assert spec.calendar().is_workday(dt.date(2026, 4, 18))   # 土


def test_day_unit_skips_the_weekday_row(make_filled):
    """日単位のファイルは見出しの下に曜日の行があるぶん、開始行がずれる。"""
    path = make_filled("day.xlsx", unit="day", end=None, months=6)
    imported = read(path)
    assert imported.spec.unit == "day"
    assert [r.name for r in imported.rows][:2] == ["要件定義", "基本設計"]


# ---------------------------------------------------------------- 列の探索
def test_columns_are_found_by_their_labels(make_filled):
    """列を入れ替えても見出しで探すので読める。"""
    path = make_filled("moved.xlsx")
    book = openpyxl.load_workbook(path)
    sheet = book[SHEET_PLAN]
    sheet.insert_cols(2)                     # 表の左に列を 1 本足す
    book.save(path)

    rows = read(path).rows
    assert [r.name for r in rows][:2] == ["要件定義", "基本設計"]
    assert rows[0].start == dt.date(2026, 4, 1)


def test_plan_and_actual_are_told_apart(filled_book):
    """「開始」「日数」は予定と実績の両方にあるので、上の帯で見分ける。"""
    row = read(filled_book).rows[1]
    assert (row.start, row.end) == (dt.date(2026, 4, 15), dt.date(2026, 5, 7))
    assert (row.actual_start, row.actual_end) == (dt.date(2026, 4, 15), dt.date(2026, 5, 12))


# ---------------------------------------------------------------- 値の解釈
@pytest.mark.parametrize("written,expected", [
    (0.8, 0.8), ("80%", 0.8), (80, 0.8), ("0.25", 0.25), (1, 1.0), (150, 1.0),
])
def test_progress_accepts_several_notations(make_filled, written, expected):
    path = make_filled("p.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 1), 5, None,
         None, None, None, written, None, "", "設計", ""),
    ])
    assert read(path).rows[0].progress == expected


def test_days_written_with_a_unit_is_read(make_filled):
    path = make_filled("d.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 1), "5 日", None,
         None, None, None, None, None, "", "設計", ""),
    ])
    assert read(path).rows[0].days == 5


def test_unreadable_values_are_reported_not_fatal(make_filled):
    path = make_filled("bad.xlsx")
    book = openpyxl.load_workbook(path)
    book[SHEET_PLAN]["F6"] = "来週くらい"
    book.save(path)

    imported = read(path)
    assert len(imported.rows) == 4                 # 読めた行は残る
    assert any("6 行目" in w for w in imported.warnings)


# ---------------------------------------------------------------- 異常系
def test_a_blank_workbook_has_no_rows(tmp_path):
    path = tmp_path / "blank.xlsx"
    write(build(dt.date(2026, 4, 1), months=3, rows=10), path)
    assert read(path).rows == []


def test_a_sheet_without_the_item_column_is_rejected(tmp_path):
    book = openpyxl.Workbook()
    book.active["A1"] = "なにか別の表"
    path = tmp_path / "other.xlsx"
    book.save(path)
    with pytest.raises(SpecError, match="見出し行が見つかりません"):
        read(path)


def test_a_non_excel_file_is_rejected(tmp_path):
    path = tmp_path / "not.xlsx"
    path.write_bytes(b"this is not a workbook")
    with pytest.raises(SpecError, match="Excel として読めません"):
        read(path)


# ---------------------------------------------------------------- 導出
def test_planned_days_are_counted_from_the_dates(make_filled):
    """予定の日数は、書かれた値ではなく開始日と終了日から数える。"""
    path = make_filled("d.xlsx", rows=[
        # 4/1〜4/14 は稼働日 10 日だが、わざと 99 と書いておく
        ("開発", "", "1", "A", dt.date(2026, 4, 1), 99, dt.date(2026, 4, 14),
         None, None, None, None, None, "", "設計", ""),
    ])
    row = read(path).rows[0]
    assert row.days == 10
    assert "days" in row.derived


def test_actual_days_are_counted_from_the_dates(make_filled):
    path = make_filled("ad.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 1), 10, dt.date(2026, 4, 14),
         dt.date(2026, 4, 1), 99, dt.date(2026, 4, 10), None, None, "", "設計", ""),
    ])
    row = read(path).rows[0]
    assert row.actual_days == 8          # 4/1〜4/10 の稼働日
    assert "actual_days" in row.derived


def test_an_actual_end_date_means_one_hundred_percent(make_filled):
    """実績の終了日が入っていれば、書かれた進捗より優先して 100% とする。"""
    path = make_filled("p.xlsx", rows=[
        ("開発", "", "1", "済", dt.date(2026, 4, 1), 10, dt.date(2026, 4, 14),
         dt.date(2026, 4, 1), None, dt.date(2026, 4, 10), 0.3, None, "", "設計", ""),
        ("開発", "", "2", "途中", dt.date(2026, 4, 1), 10, dt.date(2026, 4, 14),
         dt.date(2026, 4, 1), 5, None, 0.3, None, "", "設計", ""),
    ])
    done, running = read(path).rows
    assert done.progress == 1.0
    assert "progress" in done.derived
    assert running.progress == 0.3       # 終了日が無ければ触らない
    assert "progress" not in running.derived


def test_a_derived_actual_end_does_not_mean_finished(make_filled):
    """実績日数から終了日を補った行は、完了とはみなさない。"""
    path = make_filled("r.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 1), 10, dt.date(2026, 4, 14),
         dt.date(2026, 4, 1), 5, None, 0.5, None, "", "設計", ""),
    ])
    row = read(path).rows[0]
    assert row.actual_end == dt.date(2026, 4, 7)     # バーを描くために補う
    assert "actual_end" in row.derived
    assert row.progress == 0.5                       # 完了扱いにはしない


def test_a_missing_end_date_is_derived_from_the_days(make_filled):
    path = make_filled("e.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 1), 10, None,
         None, None, None, None, None, "", "設計", ""),
    ])
    row = read(path).rows[0]
    assert row.end == dt.date(2026, 4, 14)
    assert "end" in row.derived


def test_days_that_already_match_are_not_marked_as_derived(make_filled):
    path = make_filled("m.xlsx", rows=[
        ("開発", "", "1", "A", dt.date(2026, 4, 1), 10, dt.date(2026, 4, 14),
         None, None, None, None, None, "", "設計", ""),
    ])
    assert read(path).rows[0].derived == set()


def test_the_working_calendar_is_used_for_counting(make_filled):
    """稼働曜日の設定が数え方に効く (土曜も稼働にすると日数が増える)。"""
    weekdays = ["mon", "tue", "wed", "thu", "fri", "sat"]
    rows = [("開発", "", "1", "A", dt.date(2026, 4, 1), None, dt.date(2026, 4, 14),
             None, None, None, None, None, "", "設計", "")]
    assert read(make_filled("w1.xlsx", rows=rows)).rows[0].days == 10
    assert read(make_filled("w2.xlsx", rows=rows, workdays=weekdays)).rows[0].days == 12
