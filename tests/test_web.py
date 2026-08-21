import datetime as dt
import io
import zipfile

import openpyxl
import pytest

pytest.importorskip("fastapi", reason="Web アプリの依存 (pip install 'wbsgen[web]')")
from fastapi.testclient import TestClient  # noqa: E402

from wbsgen.web.app import app  # noqa: E402


@pytest.fixture
def client():
    with TestClient(app) as c:
        yield c


SPEC = {"start": "2026-02-01", "end": "2027-01-31", "unit": "week", "rows": 40}


# ---------------------------------------------------------------- 画面
def test_index_serves_the_app_shell(client):
    response = client.get("/")
    assert response.status_code == 200
    assert "WBS ジェネレータ" in response.text
    assert "/static/app.js" in response.text


@pytest.mark.parametrize("path", ["/static/app.js", "/static/style.css"])
def test_static_assets_are_served(client, path):
    assert client.get(path).status_code == 200


def test_meta_offers_the_choices(client):
    body = client.get("/api/meta").json()
    assert [u["value"] for u in body["units"]] == ["day", "week", "month"]
    assert [w["label"] for w in body["weekdays"]] == list("月火水木金土日")
    assert body["default_rows"] == 40
    # 既定は年度 (4 月始まり)
    assert body["suggested"]["start"].endswith("-04-01")
    assert body["suggested"]["end"].endswith("-03-31")


# ---------------------------------------------------------------- preview
def test_preview_returns_the_timeline(client):
    body = client.post("/api/preview", json=SPEC).json()
    assert body["spec"]["period_days"] == 365
    assert body["spec"]["end"] == "2027-01-31"
    timeline = body["timeline"]
    assert len(timeline["columns"]) == 53
    assert [c["label"] for c in timeline["columns"][:3]] == ["2/1", "2/8", "2/15"]
    assert [b["label"] for b in timeline["bands"]][:3] == ["2月", "3月", "4月"]
    assert [b["start"] for b in timeline["bands"]][:3] == [0, 4, 9]


def test_preview_marks_rest_columns_in_day_view(client):
    body = client.post("/api/preview",
                       json={**SPEC, "start": "2026-04-01", "months": 1,
                             "end": None, "unit": "day"}).json()
    columns = {c["start"]: c for c in body["timeline"]["columns"]}
    assert columns["2026-04-04"]["rest"] and columns["2026-04-04"]["saturday"]
    assert columns["2026-04-29"]["rest"]          # 昭和の日
    assert not columns["2026-04-30"]["rest"]
    assert columns["2026-04-01"]["weekday"] == "水"


def test_preview_lists_the_holidays(client):
    body = client.post("/api/preview", json=SPEC).json()
    assert "2026-02-11" in body["holidays"]
    assert len(body["holidays"]) == 18


def test_preview_honours_the_calendar_settings(client):
    body = client.post("/api/preview", json={
        **SPEC, "japanese_holidays": False, "holidays": ["2026-12-30"],
    }).json()
    assert body["holidays"] == ["2026-12-30"]


@pytest.mark.parametrize("payload,message", [
    ({"start": "bad"}, "start"),
    ({"start": "2026-04-01", "end": "2020-01-01"}, "終了日"),
    ({"start": "2026-04-01", "unit": "hour"}, "表示単位"),
    ({"start": "2026-04-01", "rows": 99999}, "行数"),
    ({"start": "2026-04-01", "workdays": ["someday"]}, "曜日"),
])
def test_preview_rejects_bad_input(client, payload, message):
    response = client.post("/api/preview", json=payload)
    assert response.status_code == 422
    assert message in response.json()["detail"]


# ---------------------------------------------------------------- build
def test_build_returns_a_real_workbook(client):
    response = client.post("/api/build", json={**SPEC, "members": ["設計"]})
    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml")

    book = openpyxl.load_workbook(io.BytesIO(response.content))
    assert book.sheetnames == ["スケジュール", "担当者一覧", "設定"]
    ws = book["スケジュール"]
    assert ws["S3"].value == dt.datetime(2026, 2, 1)
    assert ws["S4"].value == dt.datetime(2026, 2, 1)
    assert ws.max_row == 44
    assert book["担当者一覧"]["B4"].value == "設計"


def test_build_has_no_drawings(client):
    """空の WBS なので図形は入らない。"""
    response = client.post("/api/build", json=SPEC)
    with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
        assert not [n for n in zf.namelist() if "drawings" in n]


def test_build_filename_carries_the_title(client):
    response = client.post("/api/build", json={**SPEC, "title": "年度計画/2026"})
    disposition = response.headers["content-disposition"]
    assert "filename*=UTF-8''" in disposition
    assert "%2F" not in disposition             # 使えない文字は落としてある
    assert "20260201" in disposition


def test_build_rejects_bad_input(client):
    assert client.post("/api/build", json={"start": "bad"}).status_code == 422


# ---------------------------------------------------------------- import
def _upload(client, path, name="filled.xlsx", params=None):
    with open(path, "rb") as handle:
        return client.post("/api/import", params=params or {}, files={
            "file": (name, handle.read(),
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        })


def test_import_returns_a_chart(client, filled_book):
    body = _upload(client, filled_book).json()
    assert body["title"] == "2026年度 開発スケジュール"
    assert len(body["rows"]) == 5
    assert body["timeline"]["unit"] == "week"
    assert body["totals"]["done"] == 2
    assert body["rows"][0]["plan"]["x2"] > body["rows"][0]["plan"]["x1"]


@pytest.mark.parametrize("unit,columns", [("day", 365), ("week", 53), ("month", 12)])
def test_import_can_override_the_unit(client, filled_book, unit, columns):
    body = _upload(client, filled_book, params={"unit": unit}).json()
    assert body["timeline"]["unit"] == unit
    assert len(body["timeline"]["columns"]) == columns


def test_import_rejects_an_unknown_unit(client, filled_book):
    response = _upload(client, filled_book, params={"unit": "hour"})
    assert response.status_code == 422
    assert "表示単位" in response.json()["detail"]


def test_import_rejects_a_non_excel_file(client):
    response = client.post("/api/import", files={"file": ("notes.txt", b"hello", "text/plain")})
    assert response.status_code == 415
    assert "対応していない形式" in response.json()["detail"]


def test_import_rejects_a_broken_file(client):
    response = client.post("/api/import", files={"file": ("broken.xlsx", b"nope", "x")})
    assert response.status_code == 422
    assert "Excel として読めません" in response.json()["detail"]


def test_import_rejects_an_empty_workbook(client, tmp_path):
    """記入されていない空の WBS を読ませたら、その旨を返す。"""
    from wbsgen.blank import build as build_spec
    from wbsgen.workbook import write as write_book

    path = tmp_path / "blank.xlsx"
    write_book(build_spec(dt.date(2026, 4, 1), months=3, rows=10), path)
    response = _upload(client, path, name="blank.xlsx")
    assert response.status_code == 422
    assert "記入された行が見つかりません" in response.json()["detail"]


def test_import_reports_unreadable_rows(client, make_filled):
    import openpyxl

    path = make_filled("bad.xlsx")
    book = openpyxl.load_workbook(path)
    book["スケジュール"]["F6"] = "来週くらい"
    book.save(path)

    body = _upload(client, path, name="bad.xlsx").json()
    assert len(body["rows"]) == 4
    assert any("6 行目" in w for w in body["warnings"])
