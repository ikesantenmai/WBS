import io
import json
import zipfile
from pathlib import Path

import pytest

pytest.importorskip("fastapi", reason="Web アプリの依存 (pip install 'wbsgen[web]')")
from fastapi.testclient import TestClient  # noqa: E402

from wbsgen.web.app import app  # noqa: E402

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"


@pytest.fixture
def client():
    with TestClient(app) as c:
        yield c


@pytest.fixture
def project(client):
    return client.get("/api/template/minimal").json()


# ---------------------------------------------------------------- 画面
def test_index_serves_the_app_shell(client):
    response = client.get("/")
    assert response.status_code == 200
    assert "WBS ジェネレータ" in response.text
    assert "/static/app.js" in response.text


@pytest.mark.parametrize("path", ["/static/app.js", "/static/style.css"])
def test_static_assets_are_served(client, path):
    assert client.get(path).status_code == 200


# ---------------------------------------------------------------- メタ / 雛形
def test_meta_lists_choices(client):
    body = client.get("/api/meta").json()
    assert body["units"] == ["day", "week", "month"]
    assert "milestone" in body["kinds"]
    assert "standard" in body["templates"]


def test_template_returns_a_project(client):
    body = client.get("/api/template/standard").json()
    assert body["project"]["title"]
    assert len(body["tasks"]) > 10
    assert body["chart"]["unit"] == "week"


def test_unknown_template_is_404(client):
    assert client.get("/api/template/nope").status_code == 404


# ---------------------------------------------------------------- preview
def test_preview_resolves_dates_and_chart(client, project):
    body = client.post("/api/preview", json=project).json()
    assert body["timeline"]["unit"] == "day"
    assert body["timeline"]["columns"]
    assert len(body["rows"]) == len(project["tasks"])

    first = body["rows"][0]
    assert first["end"] is not None          # 終了日が導出されている
    assert first["plan"]["x2"] > first["plan"]["x1"]
    assert body["totals"]["tasks"] == len(project["tasks"])


def test_preview_places_bars_where_the_dates_say(client):
    payload = {
        "chart": {"start": "2026-04-01", "period_days": 30, "unit": "day",
                  "base_date": "2026-04-01"},
        "tasks": [{"name": "A", "start": "2026-04-03", "days": 2}],
    }
    row = client.post("/api/preview", json=payload).json()["rows"][0]
    # 日単位なので 1 列 = 1 日。4/3 は 3 列目 (index 2)、4/6 の終わりで閉じる
    assert row["end"] == "2026-04-06"        # 4/4,4/5 は土日
    assert row["plan"]["x1"] == 2.0
    assert row["plan"]["x2"] == 6.0


def test_preview_returns_links_and_inazuma(client):
    payload = {
        "chart": {"start": "2026-04-01", "period_days": 60, "unit": "day",
                  "base_date": "2026-04-10"},
        "tasks": [
            {"no": "1", "name": "A", "start": "2026-04-01", "days": 3, "progress": 1.0,
             "actual_start": "2026-04-01", "actual_days": 3},
            {"no": "2", "name": "B", "predecessor": "1", "days": 3},
        ],
    }
    body = client.post("/api/preview", json=payload).json()
    # A は 4/1〜4/3 (x2=3.0)、B は自動配置で翌稼働日の 4/6 (x1=5.0) から始まる
    assert body["links"] == [{"from_row": 1, "to_row": 2, "x1": 3.0, "x2": 5.0}]
    assert [p["row"] for p in body["inazuma"]] == [1, 2]
    assert body["now_x"] == 9.0


def test_preview_rejects_a_broken_definition(client):
    response = client.post("/api/preview", json={"tasks": [{"name": "A", "progress": 900}]})
    assert response.status_code == 422
    assert "progress" in response.json()["detail"]


def test_preview_reports_circular_predecessors(client):
    response = client.post("/api/preview", json={"tasks": [
        {"no": "1", "name": "A", "predecessor": "2", "days": 1},
        {"no": "2", "name": "B", "predecessor": "1", "days": 1},
    ]})
    assert response.status_code == 422
    assert "循環" in response.json()["detail"]


# ---------------------------------------------------------------- build
def test_build_returns_a_real_workbook(client, project):
    response = client.post("/api/build", json=project)
    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml")
    with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
        names = zf.namelist()
    assert "xl/worksheets/sheet1.xml" in names
    assert "xl/drawings/drawing1.xml" in names   # ガントの図形


def test_build_filename_carries_the_project_title(client, project):
    project["project"]["title"] = "テスト計画/2026"
    disposition = client.post("/api/build", json=project).headers["content-disposition"]
    assert "filename*=UTF-8''" in disposition
    assert "%2F" not in disposition             # 使えない文字は落としてある


def test_build_rejects_a_broken_definition(client):
    assert client.post("/api/build", json={"tasks": []}).status_code == 422


# ---------------------------------------------------------------- yaml
def test_export_yaml_round_trips(client, project):
    response = client.post("/api/export/yaml", json=project)
    assert response.status_code == 200
    text = response.content.decode("utf-8")
    assert "tasks:" in text

    import yaml

    again = client.post("/api/preview", json=yaml.safe_load(text))
    assert again.status_code == 200


# ---------------------------------------------------------------- import
def _upload(client, name, body, content_type="application/octet-stream"):
    return client.post("/api/import", files={"file": (name, body, content_type)})


def test_import_yaml(client):
    raw = (EXAMPLES / "sbi_web_wbs.yaml").read_bytes()
    body = _upload(client, "sbi.yaml", raw).json()
    assert len(body["tasks"]) == 128
    assert body["chart"]["base_date"] == "2026-08-07"


def test_import_csv(client):
    raw = (EXAMPLES / "tasks.csv").read_bytes()
    body = _upload(client, "tasks.csv", raw).json()
    assert body["tasks"][0]["name"] == "開発環境"
    assert [m["name"] for m in body["members"]]


def test_import_json(client, project):
    body = _upload(client, "p.json", json.dumps(project).encode()).json()
    assert len(body["tasks"]) == len(project["tasks"])


def test_import_accepts_cp932_csv(client):
    text = "項目,開始日,日数\nシフトJIS,2026-04-01,3\n"
    body = _upload(client, "sjis.csv", text.encode("cp932")).json()
    assert body["tasks"][0]["name"] == "シフトJIS"


def test_import_rejects_unsupported_extension(client):
    response = _upload(client, "notes.txt", b"hello")
    assert response.status_code == 415
    assert "対応していない形式" in response.json()["detail"]


def test_import_rejects_broken_yaml(client):
    response = _upload(client, "broken.yaml", b"tasks: [ {name: A\n")
    assert response.status_code == 422


def test_import_rejects_oversized_upload(client):
    response = _upload(client, "big.yaml", b"x" * (5 * 1024 * 1024))
    assert response.status_code == 413
