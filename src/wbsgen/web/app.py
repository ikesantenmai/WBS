"""空の WBS を作る Web アプリケーション。

期間を指定して、日程表の見た目を確かめてから Excel を書き出す。
日程表の組み立ては CLI とまったく同じ :mod:`wbsgen.timeline` を通る。
"""

from __future__ import annotations

import datetime as _dt
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


def _spec(payload: Dict[str, Any]) -> BlankWBS:
    try:
        return from_dict(payload)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))


# ----------------------------------------------------------------------
@app.get("/api/meta")
def meta() -> Dict[str, Any]:
    """画面が選択肢を組み立てるためのメタ情報。"""
    today = _dt.date.today()
    year = today.year if today.month >= 4 else today.year - 1
    return {
        "version": __version__,
        "units": [
            {"value": "day", "label": "日単位"},
            {"value": "week", "label": "週単位"},
            {"value": "month", "label": "月単位"},
        ],
        "weekdays": [{"value": k, "label": j} for k, j in zip(WEEKDAY_KEYS, WEEKDAY_JP)],
        "default_rows": DEFAULT_ROWS,
        "max_rows": MAX_ROWS,
        "today": today.isoformat(),
        # 年度 (4 月始まり) を既定の期間として提案する
        "suggested": {
            "start": f"{year}-04-01",
            "end": f"{year + 1}-03-31",
            "title": f"{year}年度 スケジュール",
        },
    }


@app.post("/api/preview")
def preview(payload: Dict[str, Any] = Body(...)) -> Dict[str, Any]:
    """日程表の見出しと寸法を返す (画面のプレビュー用)。"""
    spec = _spec(payload)
    calendar = spec.calendar()
    timeline = Timeline(spec.start, spec.period_days, spec.unit, calendar)
    return {
        "spec": to_dict(spec),
        "timeline": {
            "unit": timeline.unit,
            "start": timeline.start.isoformat(),
            "end": timeline.end.isoformat(),
            "columns": [
                {
                    "label": timeline.column_label(col),
                    "weekday": timeline.weekday_label(col),
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


@app.post("/api/build")
def build_workbook(payload: Dict[str, Any] = Body(...)) -> Response:
    """Excel を生成して返す。"""
    spec = _spec(payload)
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


def _filename(spec: BlankWBS) -> str:
    stem = "".join(c for c in spec.title if c not in '\\/:*?"<>|').strip()
    return quote(f"{stem or 'wbs'}_{spec.start:%Y%m%d}.xlsx")


# ----------------------------------------------------------------------
@app.post("/api/import")
async def import_workbook(
    file: UploadFile = File(...),
    unit: Optional[str] = Query(None, description="表示単位を上書きする (day/week/month)"),
) -> Dict[str, Any]:
    """記入済みの Excel を読み込み、ガントチャートの描画モデルを返す。

    ``unit`` を渡すと、ファイルに書かれた表示単位より優先する
    (同じ内容を日/週/月で見比べるため)。
    """
    raw = await file.read()
    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="ファイルが大きすぎます (上限 8MB)")

    name = Path(file.filename or "wbs.xlsx").name
    if Path(name).suffix.lower() not in (".xlsx", ".xlsm"):
        raise HTTPException(
            status_code=415,
            detail=f"対応していない形式です: {Path(name).suffix or '(拡張子なし)'}"
                   " — .xlsx を指定してください",
        )
    try:
        imported = read_workbook(io.BytesIO(raw), filename=name)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    if unit:
        if unit not in VALID_UNITS:
            raise HTTPException(status_code=422, detail=f"表示単位が不正です: {unit}")
        imported.spec.unit = unit

    if not imported.rows:
        raise HTTPException(
            status_code=422,
            detail="記入された行が見つかりません。項目と日付を入れてから読み込んでください。",
        )
    return build_chart(imported)


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((STATIC_DIR / "index.html").read_text(encoding="utf-8"))


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


# ----------------------------------------------------------------------
def serve(host: str = "127.0.0.1", port: int = 8000, reload: bool = False) -> None:
    """開発用サーバを起動する。"""
    import uvicorn

    uvicorn.run("wbsgen.web.app:app" if reload else app,
                host=host, port=port, reload=reload)
