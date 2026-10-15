"""旧形式 (.xls) との相互変換。

openpyxl は .xlsx / .xlsm しか扱えないので、.xls は LibreOffice を
ヘッドレスで呼び出して .xlsx に直して読み込み、書き出しのときに .xls へ
戻す。LibreOffice が無い環境では :class:`ConvertError` にする。

変換のたびに数秒かかるので、同じ中身は結果を覚えておく。
"""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
from collections import OrderedDict
from pathlib import Path
from typing import Optional

#: 変換 1 回の上限 (秒)
TIMEOUT_SECONDS = 120

#: 覚えておく変換結果の数
CACHE_SIZE = 8

FILTERS = {"xlsx": "xlsx:Calc MS Excel 2007 XML", "xls": "xls:MS Excel 97"}

_cache: "OrderedDict[tuple, bytes]" = OrderedDict()


class ConvertError(Exception):
    """変換できなかった。``reason`` は ``missing`` / ``failed``。"""

    def __init__(self, reason: str, detail: str = ""):
        super().__init__(detail or reason)
        self.reason = reason


def office_command() -> Optional[str]:
    """LibreOffice の実行ファイル。見つからなければ ``None``。"""
    return shutil.which("soffice") or shutil.which("libreoffice")


def available() -> bool:
    return office_command() is not None


def to_xlsx(raw: bytes) -> bytes:
    """.xls の中身を .xlsx に変換する。"""
    return _convert(raw, "xls", "xlsx")


def to_xls(raw: bytes) -> bytes:
    """.xlsx の中身を .xls に変換する。"""
    return _convert(raw, "xlsx", "xls")


def _convert(raw: bytes, source: str, target: str) -> bytes:
    key = (hashlib.sha256(raw).hexdigest(), target)
    if key in _cache:
        _cache.move_to_end(key)
        return _cache[key]

    command = office_command()
    if command is None:
        raise ConvertError("missing")

    with tempfile.TemporaryDirectory() as tmp:
        folder = Path(tmp)
        origin = folder / f"book.{source}"
        origin.write_bytes(raw)
        out = folder / "out"
        out.mkdir()
        try:
            # 変換のたびに専用の設定置き場を使う (同時に呼ばれても邪魔し合わない)
            subprocess.run(
                [command, f"-env:UserInstallation=file://{folder / 'profile'}",
                 "--headless", "--convert-to", FILTERS[target],
                 "--outdir", str(out), str(origin)],
                check=True, capture_output=True, timeout=TIMEOUT_SECONDS)
        except (subprocess.SubprocessError, OSError) as exc:
            raise ConvertError("failed", str(exc))
        produced = out / f"book.{target}"
        if not produced.exists():
            raise ConvertError("failed", "no output")
        data = produced.read_bytes()

    _cache[key] = data
    while len(_cache) > CACHE_SIZE:
        _cache.popitem(last=False)
    return data
