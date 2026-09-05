"""空の WBS を作る Web アプリケーション。

期間を指定して、日程表の見た目を確かめてから Excel を書き出す。
日程表の組み立ては CLI とまったく同じ :mod:`wbsgen.timeline` を通る。
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import io
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import quote

from fastapi import Body, FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from .. import __version__, workbook
from ..blank import DEFAULT_ROWS, MAX_ROWS, BlankWBS, SpecError, from_dict, to_dict
from ..chart import build as build_chart
from ..i18n import DEFAULT_LANGUAGE, LANGUAGES, labels as get_labels, message, normalize
from ..importer import read as read_workbook
from ..timeline import VALID_UNITS, Timeline
from ..workcal import WEEKDAY_JP, WEEKDAY_KEYS

STATIC_DIR = Path(__file__).parent / "static"
MAX_UPLOAD_BYTES = 8 * 1024 * 1024

app = FastAPI(
    title="wbsgen",
    description="期間を指定して、中身が空の WBS を作ります。",
    version=__version__,
)


def _spec(payload: Dict[str, Any], language: str = DEFAULT_LANGUAGE) -> BlankWBS:
    try:
        return from_dict(payload, language=language)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))


def _language(value, payload: Optional[Dict[str, Any]] = None) -> str:
    """クエリ、なければ本文から言語を決める。"""
    if value:
        return normalize(value)
    if payload:
        return normalize(payload.get("language"))
    return DEFAULT_LANGUAGE


# ----------------------------------------------------------------------
@app.get("/api/meta")
def meta(lang: Optional[str] = Query(None, description="表示言語 (ja/en)")) -> Dict[str, Any]:
    """画面が選択肢を組み立てるためのメタ情報。"""
    language = normalize(lang)
    text = get_labels(language)
    today = _dt.date.today()
    # 年度 (4 月始まり) を既定の期間として提案する
    year = today.year if today.month >= 4 else today.year - 1
    return {
        "version": __version__,
        "language": language,
        "languages": [
            {"value": "ja", "label": "日本語"},
            {"value": "en", "label": "English"},
        ],
        "units": [{"value": key, "label": text.units[key]} for key in ("day", "week", "month")],
        "weekdays": [
            {"value": key, "label": label}
            for key, label in zip(WEEKDAY_KEYS, text.weekdays)
        ],
        "default_rows": DEFAULT_ROWS,
        "max_rows": MAX_ROWS,
        "today": today.isoformat(),
        "suggested": {
            "start": f"{year}-04-01",
            "end": f"{year + 1}-03-31",
            "title": (f"{year}年度 スケジュール" if language == "ja"
                      else f"FY{year} Schedule"),
        },
    }


@app.post("/api/preview")
def preview(
    payload: Dict[str, Any] = Body(...),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
) -> Dict[str, Any]:
    """日程表の見出しと寸法を返す (画面のプレビュー用)。"""
    language = _language(lang, payload)
    spec = _spec({**payload, "language": language}, language)
    calendar = spec.calendar()
    timeline = Timeline(spec.start, spec.period_days, spec.unit, calendar, spec.language)
    return {
        "spec": to_dict(spec),
        "timeline": {
            "unit": timeline.unit,
            "start": timeline.start.isoformat(),
            "end": timeline.end.isoformat(),
            "columns": [
                {
                    "label": timeline.column_label(col),
                    "weekday": _weekday(timeline, spec, col),
                    "start": col.start.isoformat(),
                    "rest": timeline.is_rest_column(col),
                    "saturday": timeline.is_saturday_column(col),
                }
                for col in timeline.columns
            ],
            # 上段 (週表示なら月)。区切りが変わる列にだけ値が入る。
            "bands": [
                {"start": index, "span": span, "label": label}
                for index, span, _day, label in timeline.header_top()
            ],
        },
        "holidays": [
            d.isoformat()
            for d in calendar.holidays_between(timeline.start, timeline.end)
        ],
    }


def _weekday(timeline, spec: BlankWBS, col) -> str:
    """曜日の見出し (言語ごと)。日単位以外は空。"""
    if timeline.unit != "day":
        return ""
    return get_labels(spec.language).weekdays[col.start.weekday()]


@app.post("/api/build")
def build_workbook(
    payload: Dict[str, Any] = Body(...),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
) -> Response:
    """Excel を生成して返す。"""
    language = _language(lang, payload)
    spec = _spec({**payload, "language": language}, language)
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "wbs.xlsx"
        workbook.write(spec, path)
        data = path.read_bytes()

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":
                f"attachment; filename=wbs.xlsx; filename*=UTF-8''{_filename(spec)}",
        },
    )


def _chart_suffix(spec: BlankWBS) -> str:
    return "_ガント" if spec.language == "ja" else "_gantt"


def _filename(spec: BlankWBS, suffix: str = "") -> str:
    stem = "".join(c for c in spec.title if c not in '\\/:*?"<>|').strip()
    return quote(f"{stem or 'wbs'}{suffix}_{spec.start:%Y%m%d}.xlsx")


# ----------------------------------------------------------------------
@app.post("/api/import")
async def import_workbook(
    file: UploadFile = File(...),
    unit: Optional[str] = Query(None, description="表示単位を上書きする (day/week/month)"),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
) -> Dict[str, Any]:
    """記入済みの Excel を読み込み、ガントチャートの描画モデルを返す。

    ``unit`` を渡すと、ファイルに書かれた表示単位より優先する
    (同じ内容を日/週/月で見比べるため)。
    """
    return build_chart(await _read_upload(file, unit, normalize(lang)))


async def _read_upload(file: UploadFile, unit: Optional[str], language: str):
    """アップロードされた Excel を読み込む (import / export で共通)。"""
    raw = await file.read()
    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail=message(
            language, "upload_too_large", limit=MAX_UPLOAD_BYTES // (1024 * 1024)))

    name = Path(file.filename or "wbs.xlsx").name
    suffix = Path(name).suffix.lower()
    if suffix not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=415, detail=message(
            language, "unsupported_format",
            suffix=suffix or message(language, "no_suffix")))
    try:
        imported = read_workbook(io.BytesIO(raw), filename=name, language=language)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    if unit:
        if unit not in VALID_UNITS:
            raise HTTPException(status_code=422, detail=message(
                language, "bad_unit", value=unit, choices="/".join(VALID_UNITS)))
        imported.spec.unit = unit

    if not imported.rows:
        raise HTTPException(status_code=422, detail=message(language, "no_rows"))
    return imported


@app.post("/api/export")
async def export_workbook(
    file: UploadFile = File(...),
    unit: Optional[str] = Query(None, description="表示単位を上書きする (day/week/month)"),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
) -> Response:
    """記入済みの Excel を読み込み、ガントチャートを描き込んで返す。

    画面に出しているのと同じバーを、浮動図形として書き込む。
    """
    imported = await _read_upload(file, unit, normalize(lang))
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "wbs.xlsx"
        workbook.export(imported, path)
        data = path.read_bytes()

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":
                f"attachment; filename=wbs.xlsx;"
                f" filename*=UTF-8''{_filename(imported.spec, _chart_suffix(imported.spec))}",
        },
    )


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    """トップページ。

    JavaScript と CSS の読み込みに、**中身から作った印**を付ける。
    版番号だと上げ忘れたときに前の版が使われ続けるので、中身が変われば
    必ず変わるものにしてある。ページ自体は毎回取り直させる。
    """
    page = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    for asset in ("app.js", "style.css"):
        page = page.replace(f"/static/{asset}", f"/static/{asset}?v={_stamp(asset)}")
    return HTMLResponse(page, headers={
        "Cache-Control": "no-cache, no-store, must-revalidate",
    })


def _stamp(asset: str) -> str:
    """静的ファイルの中身から作る短い印 (キャッシュを切り替えるため)。"""
    try:
        digest = hashlib.sha256((STATIC_DIR / asset).read_bytes()).hexdigest()
    except OSError:
        return __version__
    return digest[:8]


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ----------------------------------------------------------------------
def serve(host: str = "127.0.0.1", port: int = 8000, reload: bool = False) -> None:
    """開発用サーバを起動する。"""
    import uvicorn

    uvicorn.run("wbsgen.web.app:app" if reload else app,
                host=host, port=port, reload=reload)
