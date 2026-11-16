"""「開発マイルストーン」(km) 形式の .xls を書き出す。

同梱の空テンプレート (ActiveX のボタン、名前、マクロ入り) に WBS の中身を
書き込み、アドイン ``km_module.xla`` と一緒に zip にして返す。テンプレートの
埋め込みは LibreOffice (python3-uno) で行う。

公開サイトでだけ使えるよう、環境変数 ``WBSGEN_KM`` (1 / true / on) で有効にする。
"""

from __future__ import annotations

import atexit
import json
import os
import shutil
import subprocess
import signal
import tempfile
import threading
import time
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional

from .. import convert

HERE = Path(__file__).parent
TEMPLATE = HERE / "blank.xls"
ADDIN = HERE / "km_module.xla"
SCRIPT = HERE / "fill_uno.py"
VBA_DIR = HERE / "vba"
TIMEOUT_SECONDS = 420          # 無料枠など遅い環境では LibreOffice の起動だけで時間がかかる


PIPE = f"wbsgen_km_{os.getpid()}"
_lock = threading.Lock()
_office: Optional[subprocess.Popen] = None


def _stop_office() -> None:
    global _office
    if _office is not None:
        _kill(_office)
        _office = None


def _ensure_office(soffice: str) -> None:
    """LibreOffice を常駐させる (呼び出しのたびに起動すると、遅い環境で時間切れになる)。"""
    global _office
    if _office is not None and _office.poll() is None:
        return
    profile = Path(tempfile.gettempdir()) / f"wbsgen_km_profile_{os.getpid()}"
    _office = subprocess.Popen(
        [soffice, "--headless", "--norestore", "--nologo",
         f"--accept=pipe,name={PIPE};urp;", f"-env:UserInstallation=file://{profile}"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)


def warm() -> None:
    """サーバーの起動時に呼ぶ。最初のダウンロードを待たせないよう、先に LibreOffice を起動しておく。"""
    soffice = convert.office_command()
    if enabled() and soffice:
        with _lock:
            _ensure_office(soffice)


atexit.register(lambda: _stop_office())


class KMError(Exception):
    pass


def enabled() -> bool:
    """この環境で km 形式を出せるか (環境変数で有効にし、道具が揃っているときだけ)。"""
    flag = os.environ.get("WBSGEN_KM", "").strip().lower() in ("1", "true", "on", "yes")
    return flag and TEMPLATE.exists() and ADDIN.exists() and available()


def available() -> bool:
    return convert.available() and uno_python() is not None


def uno_python() -> Optional[str]:
    """``import uno`` できる python。環境変数 ``WBSGEN_UNO_PYTHON``、次にシステムの python3。"""
    candidates = [os.environ.get("WBSGEN_UNO_PYTHON"), "/usr/bin/python3",
                  shutil.which("python3")]
    for path in candidates:
        if not path or not os.access(path, os.X_OK):
            continue
        try:
            subprocess.run([path, "-I", "-c", "import uno"], check=True,
                           capture_output=True, timeout=30)
            return path
        except (subprocess.SubprocessError, OSError):
            continue
    return None


def _row(row) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for key in ("group", "subgroup", "no", "name", "days", "actual_days", "delay",
                "progress", "effort", "predecessor", "member", "status"):
        value = getattr(row, key, None)
        if value not in (None, ""):
            out[key] = value
    for key in ("start", "end", "actual_start", "actual_end"):
        value = getattr(row, key, None)
        if value:
            out[key] = value.isoformat()
    return out


def _kill(child) -> None:
    try:
        os.killpg(child.pid, signal.SIGKILL)
    except (OSError, ProcessLookupError):
        pass
    child.wait()


def build(spec, rows: List[Any], stem: str) -> bytes:
    """テンプレートに書き込んだ .xls と ``km_module.xla`` を zip にして返す。"""
    python, soffice = uno_python(), convert.office_command()
    if python is None or soffice is None:
        raise KMError("missing")
    payload = {
        "soffice": soffice,
        "title": spec.title,
        "start": spec.start.isoformat(),
        "period_days": spec.period_days,
        "unit": spec.unit,                  # day / week / month
        "members": [str(m) for m in spec.members],
        # 祝祭日 (表示期間内)。設定シートの「休日設定(個別)」に書く
        "holidays": [d.isoformat() for d in spec.calendar().holidays_between(spec.start, spec.end)],
        "rows": [_row(row) for row in rows],
        "pipe": PIPE,
    }
    with _lock, tempfile.TemporaryDirectory() as tmp:       # 同時に 1 件ずつ (メモリを抑える)
        folder = Path(tmp)
        data_file, out = folder / "data.json", folder / "out.xls"
        data_file.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        _ensure_office(soffice)
        child = subprocess.Popen(
            [python, "-I", str(SCRIPT), str(TEMPLATE), str(out), str(data_file), str(VBA_DIR)],
            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, start_new_session=True)
        try:
            _, stderr = child.communicate(timeout=TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            _kill(child)
            _stop_office()                 # 固まった LibreOffice は捨てて、次は起動し直す
            raise KMError(f"timeout ({TIMEOUT_SECONDS}s)")
        if child.returncode != 0:
            _stop_office()
            detail = (stderr or b"").decode("utf-8", "replace").strip().splitlines()
            raise KMError("; ".join(detail[-3:]) or f"exit {child.returncode}")
        if not out.exists():
            raise KMError("no output")
        book = out.read_bytes()
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{stem}.xls", book)
        archive.write(ADDIN, "km_module.xla")
    return buffer.getvalue()
