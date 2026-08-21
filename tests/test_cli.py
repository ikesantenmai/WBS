import datetime as dt

import openpyxl
import pytest

from wbsgen.cli import main


def test_new_writes_a_workbook(tmp_path, capsys):
    output = tmp_path / "wbs.xlsx"
    assert main(["new", "--start", "2026-04-01", "--end", "2027-03-31",
                 "--rows", "50", "--title", "年度計画", "-o", str(output)]) == 0
    out = capsys.readouterr().out
    assert "2026-04-01 〜 2027-03-31 (365 日)" in out
    assert "50 行" in out

    ws = openpyxl.load_workbook(output)["スケジュール"]
    assert ws["B1"].value == "年度計画"
    assert ws.max_row == 54
    assert ws["S4"].value == dt.datetime(2026, 4, 1)


def test_new_defaults_to_twelve_months(tmp_path, capsys):
    assert main(["new", "--start", "2026-04-01", "-o", str(tmp_path / "a.xlsx")]) == 0
    assert "2026-04-01 〜 2027-03-31" in capsys.readouterr().out


@pytest.mark.parametrize("args,expected_columns", [
    (["--months", "3"], 13),                     # 週単位: 91 日 → 13 列
    (["--months", "3", "--unit", "day"], 91),
    (["--months", "3", "--unit", "month"], 3),
])
def test_period_and_unit_options(tmp_path, capsys, args, expected_columns):
    assert main(["new", "--start", "2026-04-01", *args,
                 "-o", str(tmp_path / "u.xlsx")]) == 0
    assert f"{expected_columns} 列" in capsys.readouterr().out


def test_members_are_listed(tmp_path):
    output = tmp_path / "m.xlsx"
    assert main(["new", "--start", "2026-04-01", "--member", "設計",
                 "--member", "製造", "-o", str(output)]) == 0
    ws = openpyxl.load_workbook(output)["担当者一覧"]
    assert [ws[f"B{r}"].value for r in (4, 5)] == ["設計", "製造"]


def test_workdays_and_holiday_options(tmp_path):
    output = tmp_path / "cal.xlsx"
    assert main(["new", "--start", "2026-04-01", "--months", "3", "--unit", "day",
                 "--workdays", "月,火,水,木,金,土", "--holiday", "2026-06-15",
                 "--no-jp-holidays", "-o", str(output)]) == 0
    ws = openpyxl.load_workbook(output)["設定"]
    values = {ws[f"B{r}"].value: ws[f"C{r}"].value for r in range(3, 20)}
    assert values["土曜日"] == "出"
    assert values["日曜日"] == "休"
    assert ws["E3"].value == dt.datetime(2026, 6, 15)    # 追加した休業日だけ
    assert ws["E4"].value is None                        # 祝日は入らない


def test_dates_accept_slashes(tmp_path, capsys):
    assert main(["new", "--start", "2026/04/01", "--months", "1",
                 "-o", str(tmp_path / "s.xlsx")]) == 0
    assert "2026-04-01" in capsys.readouterr().out


@pytest.mark.parametrize("args,message", [
    (["--start", "4月1日"], "YYYY-MM-DD"),
    (["--start", "2026-13-45"], "YYYY-MM-DD"),
    (["--start", "2026-04-01", "--end", "2020-01-01"], "終了日"),
    (["--start", "2026-04-01", "--rows", "99999"], "行数"),
    (["--start", "2026-04-01", "--workdays", "someday"], "曜日"),
    (["--start", "2026-04-01", "--holiday", "xxx"], "YYYY-MM-DD"),
])
def test_invalid_options_return_an_error(capsys, tmp_path, args, message):
    assert main(["new", *args, "-o", str(tmp_path / "x.xlsx")]) == 2
    assert message in capsys.readouterr().err


def test_period_options_are_mutually_exclusive(tmp_path):
    with pytest.raises(SystemExit):
        main(["new", "--start", "2026-04-01", "--end", "2027-03-31", "--months", "6",
              "-o", str(tmp_path / "x.xlsx")])


def test_only_new_and_serve_are_offered(capsys):
    with pytest.raises(SystemExit):
        main(["--help"])
    out = capsys.readouterr().out
    assert "{new,serve}" in out
    assert "build" not in out


# ---------------------------------------------------------------- serve
@pytest.mark.parametrize("env,expected", [
    ({}, ("127.0.0.1", 8000)),
    ({"PORT": "10000"}, ("0.0.0.0", 10000)),          # Render などが渡してくる
    ({"PORT": "10000", "HOST": "127.0.0.1"}, ("127.0.0.1", 10000)),
    ({"PORT": "not-a-number"}, ("127.0.0.1", 8000)),  # 壊れた値は無視する
])
def test_serve_defaults_follow_the_environment(monkeypatch, env, expected):
    """PaaS は待ち受けポートを環境変数 PORT で渡してくる。

    そのときは 0.0.0.0 で待ち受けないと外部から届かないので、
    引数なしの `wbsgen serve` だけで動くようにしてある。
    """
    from wbsgen import cli

    monkeypatch.delenv("PORT", raising=False)
    monkeypatch.delenv("HOST", raising=False)
    for key, value in env.items():
        monkeypatch.setenv(key, value)

    host, port, _hosted = cli._serve_defaults()
    assert (host, port) == expected


def test_serve_arguments_win_over_the_environment(monkeypatch):
    from wbsgen import cli

    monkeypatch.setenv("PORT", "10000")
    called = {}
    monkeypatch.setattr("wbsgen.web.serve",
                        lambda **kw: called.update(kw))
    assert cli.main(["serve", "--host", "192.168.0.5", "--port", "9999"]) == 0
    assert called == {"host": "192.168.0.5", "port": 9999, "reload": False}


def test_serve_needs_no_arguments_on_a_paas(monkeypatch):
    from wbsgen import cli

    monkeypatch.setenv("PORT", "10000")
    called = {}
    monkeypatch.setattr("wbsgen.web.serve", lambda **kw: called.update(kw))
    assert cli.main(["serve"]) == 0
    assert called == {"host": "0.0.0.0", "port": 10000, "reload": False}
