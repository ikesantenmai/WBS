"""空の WBS を作る Web アプリケーション。

期間を指定して、日程表の見た目を確かめてから Excel を書き出す。
日程表の組み立ては CLI とまったく同じ :mod:`wbsgen.timeline` を通る。
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import io
import json
import tempfile
from pathlib import Path
from typing import Any, Dict, Optional
from urllib.parse import quote

from fastapi import Body, FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles

from .. import __version__, convert, macro, workbook
from ..blank import DEFAULT_ROWS, MAX_ROWS, BlankWBS, SpecError, from_dict, to_dict
from ..chart import build as build_chart
from ..i18n import DEFAULT_LANGUAGE, LANGUAGES, labels as get_labels, message, normalize
from ..importer import apply_edits, number_edits, read as read_workbook
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
        # .xls を扱える (LibreOffice が入っている) か。デプロイ先の確認に使う
        "xls_supported": convert.available(),
        "office": convert.diagnose(),
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
    fmt: Optional[str] = Query(None, alias="format", description="xlsm / xlsx / xls"),
    with_macro: Optional[str] = Query(
        None, alias="macro", description="(旧) 1: マクロ有効 (.xlsm) / 0: マクロなし (.xlsx)"),
) -> Response:
    """Excel を生成して返す。

    ``format`` で、マクロ有効の .xlsm / マクロなしの .xlsx / 旧形式の .xls を選ぶ。
    省略すると .xlsm (同梱のマクロは日本語のときだけ。英語は .xlsx)。
    .xls には、同梱のマクロも引き継ぐ。
    """
    language = _language(lang, payload)
    spec = _spec({**payload, "language": language}, language)
    choice = _choice(fmt, with_macro)
    bundle = None if choice == "xlsx" else macro.bundled(spec.language)
    # .xls は、マクロ入りの .xlsm (または .xlsx) を作ってから変換する
    book = "xlsm" if bundle else "xlsx"
    extension = choice if choice in ("xls", "xlsx") else book
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / f"wbs.{book}"
        workbook.write(spec, path, bundle)
        data = path.read_bytes()
    if extension == "xls":
        try:
            data = convert.to_xls(data, bool(bundle))
        except convert.ConvertError as exc:
            raise HTTPException(status_code=422, detail=message(
                language, "xls_needs_office" if exc.reason == "missing" else "xls_failed",
                reason=exc))

    media = {"xls": XLS_TYPE, "xlsm": XLSM_TYPE, "xlsx": XLSX_TYPE}[extension]
    return Response(
        content=data,
        media_type=media,
        headers={
            "Content-Disposition":
                f"attachment; filename=wbs.{extension};"
                f" filename*=UTF-8''{_filename(spec, '', extension)}",
        },
    )


def _chart_suffix(spec: BlankWBS) -> str:
    return "_ガント" if spec.language == "ja" else "_gantt"


XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLSM_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.12"
XLS_TYPE = "application/vnd.ms-excel"


def _flag(value: Optional[str]) -> Optional[bool]:
    """``macro`` の指定 (1/0, on/off, true/false)。省略や読めない値は ``None``。"""
    text = (value or "").strip().lower()
    if text in ("1", "on", "true", "yes"):
        return True
    if text in ("0", "off", "false", "no"):
        return False
    return None


FORMATS = ("xlsm", "xlsx", "xls")


def _choice(fmt: Optional[str], with_macro: Optional[str] = None) -> Optional[str]:
    """ダウンロードの形式の指定 (``xlsm`` / ``xlsx`` / ``xls``)。省略なら ``None``。

    以前の ``macro`` の指定 (1 = マクロ有効、0 = マクロなし) も受け付ける。
    """
    text = (fmt or "").strip().lower().lstrip(".")
    if text in FORMATS:
        return text
    flag = _flag(with_macro)
    return None if flag is None else ("xlsm" if flag else "xlsx")


class Output:
    """書き出すファイルの形式 (拡張子・マクロの扱い)。

    ``choice`` は、画面で選んだ形式。省略なら、読み込んだファイルの形式のまま
    (旧形式 .xls で読めば .xls、.xlsm なら .xlsm、…)。
    """

    def __init__(self, imported, choice: Optional[str]):
        self.imported = imported
        language = imported.spec.language
        self.bundle, self.keep_vba, self.macro_book = None, False, False
        if choice == "xls" or (choice is None and imported.legacy):
            self.extension, self.type = "xls", XLS_TYPE
            if not imported.legacy:               # 旧形式で読んだときは、マクロを足さない
                self.bundle = None if imported.has_vba else macro.bundled(language)
                self.keep_vba = imported.macro
            # マクロ入りの .xlsm を作ってから .xls に変換する (マクロも .xls に引き継ぐ)
            self.macro_book = bool(self.bundle) or (imported.has_vba and self.keep_vba)
            return
        if choice == "xlsx":                      # マクロなし: 読み込んだマクロも落とす
            self.extension, self.type = "xlsx", XLSX_TYPE
            return
        # 読み込んだファイルにマクロがあれば引き継ぎ、無ければ同梱のマクロを組み込む
        self.bundle = None if imported.has_vba else macro.bundled(language)
        self.keep_vba = imported.macro
        if imported.macro or self.bundle:
            self.extension, self.type = "xlsm", XLSM_TYPE
        else:
            self.extension, self.type = "xlsx", XLSX_TYPE

    def disposition(self, suffix: str = "") -> str:
        return (f"attachment; filename=wbs.{self.extension};"
                f" filename*=UTF-8''{_filename(self.imported.spec, suffix, self.extension)}")


def _filename(spec: BlankWBS, suffix: str = "", extension: str = "xlsx") -> str:
    stem = "".join(c for c in spec.title if c not in '\\/:*?"<>|').strip()
    return quote(f"{stem or 'wbs'}{suffix}_{spec.start:%Y%m%d}.{extension}")


# ----------------------------------------------------------------------
@app.post("/api/import")
async def import_workbook(
    file: UploadFile = File(...),
    unit: Optional[str] = Query(None, description="表示単位を上書きする (day/week/month)"),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
    edits: Optional[str] = Form(None, description="画面で書き換えた値 (JSON)"),
) -> Dict[str, Any]:
    """記入済みの Excel を読み込み、ガントチャートの描画モデルを返す。

    ``unit`` を渡すと、ファイルに書かれた表示単位より優先する
    (同じ内容を日/週/月で見比べるため)。
    """
    return build_chart(await _read_upload(file, unit, normalize(lang), edits))


def _parse_edits(edits: Optional[str], language: str) -> Dict[str, Any]:
    """画面から送られた編集内容 (JSON) を辞書にする。"""
    if not edits:
        return {}
    try:
        parsed = json.loads(edits)
    except ValueError:
        parsed = None
    if not isinstance(parsed, dict):
        raise HTTPException(status_code=422, detail=message(language, "bad_edits"))
    return parsed


async def _read_upload(file: UploadFile, unit: Optional[str], language: str,
                       edits: Optional[str] = None):
    """アップロードされた Excel を読み込む (import / export で共通)。

    ``edits`` があれば、画面で書き換えられた値を反映する。
    """
    raw = await file.read()
    if len(raw) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail=message(
            language, "upload_too_large", limit=MAX_UPLOAD_BYTES // (1024 * 1024)))

    name = Path(file.filename or "wbs.xlsx").name
    suffix = Path(name).suffix.lower()
    if suffix not in (".xlsx", ".xlsm", ".xls"):
        raise HTTPException(status_code=415, detail=message(
            language, "unsupported_format",
            suffix=suffix or message(language, "no_suffix")))
    converted = None
    if suffix == ".xls":
        # 旧形式は .xlsx に直して読む。書き出しで .xls に戻す。
        try:
            converted = convert.to_xlsx(raw)
        except convert.ConvertError as exc:
            raise HTTPException(
                status_code=501 if exc.reason == "missing" else 422,
                detail=message(language, "xls_needs_office" if exc.reason == "missing"
                               else "xls_failed", reason=exc))
    try:
        imported = read_workbook(io.BytesIO(converted or raw), filename=name,
                                 language=language)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    if converted:
        imported.legacy, imported.converted = True, converted

    if unit:
        if unit not in VALID_UNITS:
            raise HTTPException(status_code=422, detail=message(
                language, "bad_unit", value=unit, choices="/".join(VALID_UNITS)))
        imported.spec.unit = unit

    if not imported.rows:
        raise HTTPException(status_code=422, detail=message(language, "no_rows"))
    try:
        apply_edits(imported, _parse_edits(edits, language), language)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    return imported


@app.post("/api/export")
async def export_workbook(
    file: UploadFile = File(...),
    unit: Optional[str] = Query(None, description="表示単位を上書きする (day/week/month)"),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
    edits: Optional[str] = Form(None, description="画面で書き換えた値 (JSON)"),
    fmt: Optional[str] = Query(None, alias="format", description="xlsm / xlsx / xls"),
    with_macro: Optional[str] = Query(
        None, alias="macro", description="(旧) 1: マクロ有効 (.xlsm) / 0: マクロなし (.xlsx)"),
) -> Response:
    """記入済みの Excel を読み込み、ガントチャートを描き込んで返す。

    画面に出しているのと同じバーを、浮動図形として書き込む。``format`` で、
    マクロ有効の .xlsm / マクロなしの .xlsx / 旧形式の .xls を選ぶ。省略すると、
    読み込んだファイルの形式のまま (マクロが無ければ、同梱のマクロを組み込む)。
    """
    language = normalize(lang)
    imported = await _read_upload(file, unit, language, edits)
    output = Output(imported, _choice(fmt, with_macro))
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "wbs.xlsx"
        workbook.export(imported, path, macro=output.bundle, keep_vba=output.keep_vba)
        data = _finish(output, path.read_bytes(), language)

    return Response(
        content=data,
        media_type=output.type,
        headers={"Content-Disposition": output.disposition(_chart_suffix(imported.spec))},
    )


@app.post("/api/number")
async def number_workbook(
    file: UploadFile = File(...),
    lang: Optional[str] = Query(None, description="表示言語 (ja/en)"),
    edits: Optional[str] = Form(None, description="画面で書き換えた値 (JSON)"),
) -> Dict[str, Any]:
    """項番が空の行に振る連番を、画面の編集として返す。

    ファイルは書き換えない (書き戻すと、数式の計算結果やマクロを失うため)。
    画面はこの編集を重ねて読み込み直し、書き出しにも反映する。
    """
    language = normalize(lang)
    imported = await _read_upload(file, None, language, edits)
    try:
        numbered, count = number_edits(imported, _parse_edits(edits, language), language)
    except SpecError as exc:
        raise HTTPException(status_code=422, detail=str(exc))
    return {"edits": numbered, "count": count}


def _finish(output: "Output", data: bytes, language: str) -> bytes:
    """書き出す中身を、選んだ形式に合わせる (.xls なら変換する)。"""
    if output.extension != "xls":
        return data
    try:
        return convert.to_xls(data, output.macro_book)
    except convert.ConvertError as exc:
        raise HTTPException(status_code=422,
                            detail=message(language, "xls_failed", reason=exc))


@app.middleware("http")
async def no_store_api(request, call_next):
    """API の結果はブラウザに保存させない (古い答えが画面に残らないように)。"""
    response = await call_next(request)
    if request.url.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


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
