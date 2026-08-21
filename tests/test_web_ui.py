"""ブラウザ経由の疎通確認。

Playwright と Chromium がある環境でだけ動く。無い場合は skip する。
JavaScript が実際に動いて表とチャートを描けることを確かめるのが目的で、
細かい見た目は対象にしない。
"""

from __future__ import annotations

import contextlib
import os
import socket
import threading
import time
from pathlib import Path

import pytest

pytest.importorskip("fastapi", reason="Web アプリの依存が必要")
sync_playwright = pytest.importorskip(
    "playwright.sync_api", reason="playwright が必要"
).sync_playwright

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"

#: 環境が用意した Chromium (PLAYWRIGHT_BROWSERS_PATH 配下) を優先して探す
def _chromium_path():
    explicit = os.environ.get("WBSGEN_CHROMIUM")
    if explicit and Path(explicit).exists():
        return explicit
    root = Path(os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "/opt/pw-browsers"))
    if root.is_dir():
        for candidate in sorted(root.glob("chromium-*/chrome-linux/chrome"), reverse=True):
            return str(candidate)
    return None


@pytest.fixture(scope="module")
def server():
    """テスト用に uvicorn をバックグラウンド起動する。"""
    import uvicorn

    from wbsgen.web.app import app

    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]

    config = uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning")
    instance = uvicorn.Server(config)
    thread = threading.Thread(target=instance.run, daemon=True)
    thread.start()
    for _ in range(100):
        if instance.started:
            break
        time.sleep(0.05)
    else:
        pytest.skip("テスト用サーバを起動できませんでした")

    yield f"http://127.0.0.1:{port}/"

    instance.should_exit = True
    thread.join(timeout=5)


@pytest.fixture(scope="module")
def page(server):
    with sync_playwright() as playwright:
        launch = {"args": ["--no-sandbox"]}
        path = _chromium_path()
        if path:
            launch["executable_path"] = path
        try:
            browser = playwright.chromium.launch(**launch)
        except Exception as exc:  # noqa: BLE001 - ブラウザが無い環境では skip
            pytest.skip(f"Chromium を起動できません: {exc}")

        context = browser.new_context(viewport={"width": 1440, "height": 900},
                                      accept_downloads=True)
        view = context.new_page()
        errors = []
        view.on("pageerror", lambda e: errors.append(str(e)))
        view.on("console",
                lambda m: errors.append(m.text) if m.type == "console" else None)
        view.errors = errors
        view.goto(server, wait_until="networkidle")
        view.wait_for_selector("table.wbs tbody tr")
        yield view
        with contextlib.suppress(Exception):
            browser.close()


# ----------------------------------------------------------------------
def test_page_renders_table_and_chart(page):
    assert page.eval_on_selector_all("table.wbs tbody tr", "n => n.length") > 10
    assert page.eval_on_selector_all("svg.gantt rect", "n => n.length") > 5
    assert "行" in page.text_content("#status-chip")
    assert page.errors == []


def test_switching_to_day_unit_redraws_the_axis(page):
    weeks = page.eval_on_selector("svg.gantt", "n => n.viewBox.baseVal.width")
    page.select_option("#chart-unit", "day")
    page.wait_for_function(
        "() => document.querySelector('svg.gantt').dataset.unit === 'day'")
    labels = page.eval_on_selector_all("svg.gantt text", "n => n.map(x => x.textContent)")
    assert "月" in labels and "土" in labels          # 曜日の見出しが出る
    assert any(label.isdigit() for label in labels)   # 日付の見出しが出る

    page.select_option("#chart-unit", "week")
    page.wait_for_function(
        "() => document.querySelector('svg.gantt').dataset.unit === 'week'")
    assert page.eval_on_selector("svg.gantt", "n => n.viewBox.baseVal.width") == weeks
    assert page.errors == []


def test_editing_a_row_recomputes_the_schedule(page):
    page.click("table.wbs tbody tr:nth-child(2)")
    page.wait_for_selector("#task-dialog[open]")
    page.fill("#task-form input[name=days]", "9")
    page.click("#task-form button[type=submit]")
    page.wait_for_function(
        "() => document.querySelector("
        "'table.wbs tbody tr:nth-child(2) td:nth-child(6)').textContent.trim() === '9 日'")
    assert page.errors == []


def test_adding_a_row_grows_the_table(page):
    before = page.eval_on_selector_all("table.wbs tbody tr", "n => n.length")
    page.click("#btn-add-task")
    page.wait_for_selector("#task-dialog[open]")
    page.fill("#task-form input[name=name]", "ブラウザから追加")
    page.click("#task-form button[type=submit]")
    page.wait_for_function(
        f"() => document.querySelectorAll('table.wbs tbody tr').length === {before + 1}")
    assert page.errors == []


def test_importing_a_file_replaces_the_project(page):
    page.set_input_files("#import-file", str(EXAMPLES / "sbi_web_wbs.yaml"))
    page.wait_for_function(
        "() => document.querySelectorAll('table.wbs tbody tr').length === 128")
    assert "SBI" in page.input_value("#title")
    assert page.errors == []


def test_downloading_the_workbook(page, tmp_path):
    with page.expect_download() as download:
        page.click("#btn-build")
    saved = tmp_path / "out.xlsx"
    download.value.save_as(saved)
    assert saved.stat().st_size > 10000
    assert download.value.suggested_filename.endswith(".xlsx")
    assert page.errors == []
