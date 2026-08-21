import datetime as dt

import pytest
import yaml

from wbsgen.loader import ProjectError, from_dict, load


def test_bare_no_key_is_treated_as_item_number():
    # YAML 1.1 では素の `no:` は真偽値 False に解決される
    data = yaml.safe_load("tasks:\n  - {no: 101, name: A, start: 2026-04-01, days: 1}\n")
    project = from_dict(data)
    assert project.tasks[0].no == "101"


def test_id_is_an_alias_for_no():
    project = from_dict({"tasks": [{"id": "A-1", "name": "A", "days": 1}]})
    assert project.tasks[0].no == "A-1"


def test_group_and_subgroup_are_inherited():
    project = from_dict({"tasks": [
        {"group": "開発", "subgroup": "設計", "name": "A", "days": 1},
        {"name": "B", "days": 1},
        {"subgroup": "製造", "name": "C", "days": 1},
        {"group": "試験", "name": "D", "days": 1},
    ]})
    got = [(t.group, t.subgroup) for t in project.tasks]
    assert got == [("開発", "設計"), ("開発", "設計"), ("開発", "製造"), ("試験", "")]


@pytest.mark.parametrize("value,expected", [
    (0.8, 0.8), ("80%", 0.8), (80, 0.8), ("0.25", 0.25), (1, 1.0),
])
def test_progress_accepts_ratio_and_percent(value, expected):
    project = from_dict({"tasks": [{"name": "A", "days": 1, "progress": value}]})
    assert project.tasks[0].progress == expected


def test_progress_out_of_range_is_rejected():
    with pytest.raises(ProjectError, match="progress"):
        from_dict({"tasks": [{"name": "A", "days": 1, "progress": 120}]})


@pytest.mark.parametrize("value", ["2026-04-01", "2026/04/01", dt.date(2026, 4, 1)])
def test_dates_accept_common_formats(value):
    project = from_dict({"tasks": [{"name": "A", "days": 1, "start": value}]})
    assert project.tasks[0].start == dt.date(2026, 4, 1)


def test_invalid_date_is_rejected():
    with pytest.raises(ProjectError, match="日付"):
        from_dict({"tasks": [{"name": "A", "days": 1, "start": "4月1日"}]})


def test_unknown_kind_is_rejected():
    with pytest.raises(ProjectError, match="kind"):
        from_dict({"tasks": [{"name": "A", "kind": "bar"}]})


def test_unknown_show_flag_is_rejected():
    with pytest.raises(ProjectError, match="chart.show"):
        from_dict({"chart": {"show": {"rainbow": True}}, "tasks": [{"name": "A"}]})


def test_missing_name_is_rejected():
    with pytest.raises(ProjectError, match="name"):
        from_dict({"tasks": [{"days": 1}]})


def test_empty_tasks_is_rejected():
    with pytest.raises(ProjectError, match="tasks"):
        from_dict({"tasks": []})


def test_load_csv(tmp_path):
    path = tmp_path / "tasks.csv"
    path.write_text(
        "大項目,項番,項目,開始日,日数,担当,進捗\n"
        "開発,01,設計,2026-04-01,5,佐藤,100%\n"
        "開発,02,製造,2026-04-08,10,鈴木,50%\n",
        encoding="utf-8",
    )
    project = load(path)
    assert [t.name for t in project.tasks] == ["設計", "製造"]
    assert project.tasks[0].progress == 1.0
    assert [m.name for m in project.members] == ["佐藤", "鈴木"]


def test_csv_with_unknown_column_is_rejected(tmp_path):
    path = tmp_path / "tasks.csv"
    path.write_text("項目,予算\nA,100\n", encoding="utf-8")
    with pytest.raises(ProjectError, match="未知の列"):
        load(path)


def test_unsupported_extension_is_rejected(tmp_path):
    path = tmp_path / "tasks.txt"
    path.write_text("x", encoding="utf-8")
    with pytest.raises(ProjectError, match="拡張子"):
        load(path)
