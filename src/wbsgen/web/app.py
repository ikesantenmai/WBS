"""WBS ジェネレータの Web アプリケーション。

ブラウザ上で WBS を編集し、ガントチャートを確認しながら Excel を書き出す。
チャートの座標も日程計算も CLI とまったく同じエンジンを通るため、
画面のプレビューと生成される Excel は一致する。
"""

from __future__ import annotations

import datetime as _dt
import json
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional

from fastapi import Body, FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from .. import __version__
from ..blank import DEFAULT_ROWS, build_blank
from ..loader import ProjectError, from_csv, from_dict
from ..model import MILESTONE_SHAPES, VALID_KINDS, VALID_UNITS
from ..preview import build_preview
from ..render import WorkbookRenderer
from ..serialize import to_dict, to_yaml
from ..templates import TEMPLATE_NAMES, template_path

STATIC_DIR = Path(__file__).parent / "static"
MAX_UPLOAD_BYTES = 4 * 1024 * 1024

app = FastAPI(
    title="wbsgen",
    description="WBS / ガントチャート Excel 生成アプリケーション",
    version=__version__,
)


# ----------------------------------------------------------------------
# エラー処理
# ----------------------------------------------------------------------
def _project(payload: Dict[str, Any]):
    """リクエストボディからプロジェクトを組み立てる。"""
    try:
        return from_dict(payload)
    except ProjectError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))


def _preview(project) -> Dict[str, Any]:
    try:
        return build_preview(project)
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))


# ----------------------------------------------------------------------
# API
# ----------------------------------------------------------------------
@app.get("/api/meta")
def meta() -> Dict[str, Any]:
    """UI が選択肢を組み立てるためのメタ情報。"""
    return {
        "version": __version__,
        "units": list(VALID_UNITS),
        "kinds": list(VALID_KINDS),
        "milestone_shapes": list(MILESTONE_SHAPES),
        "templates": list(TEMPLATE_NAMES),
        "default_blank_rows": DEFAULT_ROWS,
        "today": _dt.date.today().isoformat(),
    }


@app.get("/api/template/{name}")
def template(name: str) -> Dict[str, Any]:
    """雛形をプロジェクト定義として返す。"""
    if name not in TEMPLATE_NAMES:
        raise HTTPException(status_code=404, detail=f"雛形が見つかりません: {name}")
    import yaml

    data = yaml.safe_load(template_path(name).read_text(encoding="utf-8"))
    return to_dict(_project(data))


@app.get("/api/blank")
def blank(
    start: str = Query(..., description="開始日 (YYYY-MM-DD)"),
    end: Optional[str] = Query(None, description="終了日 (YYYY-MM-DD)"),
    months: Optional[int] = Query(None, ge=1, le=120, description="期間 (月数)"),
    unit: str = Query("week", description="表示単位 (day/week/month)"),
    rows: int = Query(DEFAULT_ROWS, ge=0, le=2000, description="記入用の空行数"),
    title: Optional[str] = Query(None, description="プロジェクト名"),
) -> Dict[str, Any]:
    """期間だけを決めた、中身が空のプロジェクト定義を返す。"""
    if unit not in VALID_UNITS:
        raise HTTPException(status_code=422, detail=f"表示単位が不正です: {unit}")
    begin = _parse_date(start, "start")
    try:
        project = build_blank(
            start=begin,
            end=_parse_date(end, "end") if end else None,
            months=months,
            unit=unit,
            rows=rows,
            title=title or f"{begin.year}年 スケジュール",
        )
    except ProjectError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    return to_dict(project)


def _parse_date(text: str, field: str) -> _dt.date:
    try:
        return _dt.datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(
            status_code=422,
            detail=f"{field} は YYYY-MM-DD 形式で指定してください: {text!r}",
        )


@app.post("/api/preview")
def preview(payload: Dict[str, Any] = Body(...)) -> Dict[str, Any]:
    """日程を解決し、表とチャートの描画モデルを返す。"""
    return _preview(_project(payload))


@app.post("/api/import")
async def import_file(file: UploadFile = File(...)) -> Dict[str, Any]:
    """YAML / JSON / CSV を読み込んでプロジェクト定義に変換する。"""
    raw = await file.read()
    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="ファイルが大きすぎます (上限 4MB)")
    name = Path(file.filename or "upload").name
    suffix = Path(name).suffix.lower()
    if suffix not in (".yaml", ".yml", ".json", ".csv", ".tsv"):
        raise HTTPException(
            status_code=415,
            detail=f"対応していない形式です: {suffix or '(拡張子なし)'}"
                   " — .yaml / .json / .csv を指定してください",
        )
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        try:
            text = raw.decode("cp932")   # Excel が書き出した CSV
        except UnicodeDecodeError:
            raise HTTPException(status_code=422, detail="文字コードを判別できません (UTF-8 か CP932)")

    import yaml

    try:
        if suffix in (".csv", ".tsv"):
            with tempfile.TemporaryDirectory() as tmp:
                path = Path(tmp) / name
                path.write_text(text, encoding="utf-8")
                project = from_csv(path)
        elif suffix == ".json":
            project = from_dict(json.loads(text))
        else:
            project = from_dict(yaml.safe_load(text) or {})
    except ProjectError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    except (json.JSONDecodeError, yaml.YAMLError) as exc:
        raise HTTPException(status_code=422, detail=f"構文を解析できません: {exc}")
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=f"読み込みに失敗しました: {exc}")

    project.title = project.title or Path(name).stem
    return to_dict(project)


@app.post("/api/export/yaml")
def export_yaml(payload: Dict[str, Any] = Body(...)) -> Response:
    """定義を YAML テキストとして書き出す。"""
    text = to_yaml(_project(payload))
    return Response(
        content=text.encode("utf-8"),
        media_type="text/yaml; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="wbs.yaml"'},
    )


@app.post("/api/build")
def build(payload: Dict[str, Any] = Body(...)) -> Response:
    """Excel ブックを生成して返す。"""
    project = _project(payload)
    try:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "wbs.xlsx"
            WorkbookRenderer(project).save(path)
            data = path.read_bytes()
    except ValueError as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    filename = _filename(project.title)
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":
                f"attachment; filename=wbs.xlsx; filename*=UTF-8''{filename}",
        },
    )


def _filename(title: str) -> str:
    """ダウンロード名を URL エンコードして返す。"""
    from urllib.parse import quote

    stem = "".join(c for c in (title or "wbs") if c not in '\\/:*?"<>|').strip()
    stamp = _dt.date.today().strftime("%Y%m%d")
    return quote(f"{stem or 'wbs'}_{stamp}.xlsx")


# ----------------------------------------------------------------------
# 画面
# ----------------------------------------------------------------------
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
