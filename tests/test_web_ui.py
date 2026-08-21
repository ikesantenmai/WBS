"""ブラウザ経由の疎通確認。

Playwright と Chromium がある環境でだけ動く。無い場合は skip する。
JavaScript が実際に動いてプレビューを描き、Excel を落とせることを確かめる。
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


def _chromium_path():
    """環境が用意した Chromium を探す。"""
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
    import uvicorn

    from wbsgen.web.app import app

    with socket.socket() as probe:
        probe.bind(("127.0.0.1", 0))
        port = probe.getsockname()[1]

    instance = uvicorn.Server(
        uvicorn.Config(app, host="127.0.0.1", port=port, log_level="warning"))
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
        view.errors = errors
        view.goto(server, wait_until="networkidle")
        view.wait_for_selector("table.wbs tbody tr")
        yield view
        with contextlib.suppress(Exception):
            browser.close()


def _reset(page):
    """既定の指定 (年度・週単位) に戻し、プレビューが描き終わるまで待つ。"""
    page.select_option("#unit", "week")
    page.check("input[name=mode][value=end]")
    # 指定が通っていればダウンロードできる状態になる
    page.wait_for_function("() => !document.querySelector('#btn-build').disabled")
    page.wait_for_selector("table.wbs tbody tr")


# ----------------------------------------------------------------------
def test_preview_renders_the_sheet(page):
    _reset(page)
    months = page.eval_on_selector_all("th.month", "n => n.map(x => x.textContent)")
    assert [m for m in months if m][:3] == ["4月", "5月", "6月"]
    assert page.eval_on_selector_all("table.wbs tbody tr", "n => n.length") > 10
    assert "空行" in page.text_content("#preview-info")
    assert page.errors == []


def test_month_field_is_hidden_until_selected(page):
    _reset(page)
    assert page.is_hidden("#field-months")
    page.check("input[name=mode][value=months]")
    page.wait_for_selector("#field-months:not([hidden])")
    assert page.is_hidden("#field-end")
    page.check("input[name=mode][value=end]")
    page.wait_for_selector("#field-end:not([hidden])")
    assert page.errors == []


def test_changing_the_unit_redraws_the_axis(page):
    _reset(page)
    page.select_option("#unit", "day")
    page.wait_for_function(
        "() => document.querySelectorAll('table.wbs thead tr').length === 3")
    weekdays = page.eval_on_selector_all(
        "table.wbs thead tr:nth-child(3) th", "n => n.map(x => x.textContent)")
    assert "土" in weekdays and "日" in weekdays

    page.select_option("#unit", "week")
    page.wait_for_function(
        "() => document.querySelectorAll('table.wbs thead tr').length === 2")
    assert page.errors == []


def test_changing_the_row_count_changes_the_preview(page):
    _reset(page)
    page.fill("#rows", "7")
    page.dispatch_event("#rows", "change")
    page.wait_for_function(
        "() => document.querySelectorAll('table.wbs tbody tr').length === 7")
    assert "空行 7 行" in page.text_content("#preview-info")
    page.fill("#rows", "40")
    page.dispatch_event("#rows", "change")
    assert page.errors == []


def test_extra_holidays_are_counted(page):
    _reset(page)
    before = page.text_content("#holiday-note")
    page.fill("#holidays", "2026-12-30, 2026-12-31")
    page.wait_for_function(
        f"() => document.querySelector('#holiday-note').textContent !== {before!r}")
    assert "祝日・休業日" in page.text_content("#holiday-note")
    page.fill("#holidays", "")
    assert page.errors == []


def test_invalid_period_shows_a_message_and_blocks_download(page):
    _reset(page)
    page.fill("#end", "2020-01-01")
    page.dispatch_event("#end", "change")
    page.wait_for_selector("#banner:not([hidden])")
    assert "終了日" in page.text_content("#banner")
    assert page.is_disabled("#btn-build")

    page.fill("#end", "2027-03-31")
    page.dispatch_event("#end", "change")
    page.wait_for_function("() => !document.querySelector('#btn-build').disabled")
    assert page.errors == []


def test_downloading_the_workbook(page, tmp_path):
    _reset(page)
    with page.expect_download() as download:
        page.click("#btn-build")
    saved = tmp_path / "out.xlsx"
    download.value.save_as(saved)
    assert saved.stat().st_size > 8000
    assert download.value.suggested_filename.endswith(".xlsx")
    assert page.errors == []
