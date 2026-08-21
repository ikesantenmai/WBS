"""テスト共通の道具。"""

import datetime as dt

import openpyxl
import pytest

from wbsgen import workbook
from wbsgen.blank import build

D = dt.date

#: 人が記入した想定の中身
#: (大項目, 中項目, 項番, 項目, 予定開始, 日数, 予定終了, 実績開始, 実績日数,
#:  実績終了, 進捗, 工数, 先行, 担当, 状態)
FILLED_ROWS = [
    ("開発", "要件", "101", "要件定義", D(2026, 4, 1), 10, D(2026, 4, 14),
     D(2026, 4, 1), 10, D(2026, 4, 14), 1.0, 10, "", "設計", "完了"),
    ("", "", "102", "基本設計", D(2026, 4, 15), 15, D(2026, 5, 7),
     D(2026, 4, 15), 18, D(2026, 5, 12), 1.0, 15, "101", "設計", "完了"),
    ("", "製造", "201", "詳細設計", D(2026, 5, 8), 10, D(2026, 5, 22),
     D(2026, 5, 13), 10, None, 0.6, 10, "102", "製造", "実行中"),
    ("", "", "202", "コーディング", D(2026, 5, 25), 20, D(2026, 6, 19),
     D(2026, 6, 1), 15, None, 0.35, 40, "201", "製造", "実行中"),
    # 終了日を書いていない行 (日数から補われる)
    ("テスト", "結合", "301", "結合テスト", D(2026, 7, 27), 15, None,
     None, None, None, 0.0, 30, "202", "テスト", "-"),
]

COLUMNS = ("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "M", "N", "O", "Q", "R")


def write_filled(path, rows=FILLED_ROWS, **spec_kwargs):
    """空の WBS を作り、``rows`` を書き込んだファイルを作る。"""
    options = dict(start=D(2026, 4, 1), end=D(2027, 3, 31), rows=30,
                   title="2026年度 開発スケジュール", members=["設計", "製造", "テスト"])
    options.update(spec_kwargs)
    spec = build(**options)
    workbook.write(spec, path)

    book = openpyxl.load_workbook(path)
    sheet = book[workbook.SHEET_PLAN]
    # 日単位表示では見出しの下に曜日の行が入るぶん、記入欄が 1 行下がる
    first_row = 6 if spec.unit == "day" else 5
    for index, values in enumerate(rows):
        row = first_row + index
        for letter, value in zip(COLUMNS, values):
            if value not in (None, ""):
                sheet[f"{letter}{row}"] = value
    book.save(path)
    return path


@pytest.fixture
def make_filled(tmp_path):
    """記入済みの WBS ファイルを作る。

    ``make_filled(name="a.xlsx", rows=[...], unit="day")`` のように、
    中身も用紙の指定も差し替えられる。
    """
    def factory(name="filled.xlsx", **kwargs):
        return write_filled(tmp_path / name, **kwargs)
    return factory


@pytest.fixture
def filled_book(make_filled):
    """既定の中身を書き込んだ WBS ファイル。"""
    return make_filled()
