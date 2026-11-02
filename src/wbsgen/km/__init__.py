"""「開発マイルストーン」(km) 形式の .xls を書き出す。

同梱の空テンプレート (ActiveX のボタン、名前、マクロ入り) に WBS の中身を
書き込み、アドイン ``km_module.xla`` と一緒に zip にして返す。テンプレートの
埋め込みは LibreOffice (python3-uno) で行う。

公開サイトでだけ使えるよう、環境変数 ``WBSGEN_KM`` (1 / true / on) で有効にする。
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Optional

from .. import convert

HERE = Path(__file__).parent
TEMPLATE = HERE / "blank.xls"
ADDIN = HERE / "km_module.xla"
SCRIPT = HERE / "fill_uno.py"
TIMEOUT_SECONDS = 180


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
        "members": [str(m) for m in spec.members],
        "rows": [_row(row) for row in rows],
        "port": 2100 + os.getpid() % 800,
    }
    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        data_file, out = folder / "data.json", folder / "out.xls"
        data_file.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        try:
            subprocess.run([python, "-I", str(SCRIPT), str(TEMPLATE), str(out), str(data_file)],
                           check=True, capture_output=True, timeout=TIMEOUT_SECONDS)
        except (subprocess.SubprocessError, OSError) as exc:
            raise KMError(str(exc))
        if not out.exists():
            raise KMError("no output")
        book = out.read_bytes()
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(f"{stem}.xls", book)
        archive.write(ADDIN, "km_module.xla")
    return buffer.getvalue()
