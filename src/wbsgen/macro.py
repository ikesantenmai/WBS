"""同梱の VBA マクロを、書き出した Excel に組み込む。

``macros/vbaProject.bin`` (「スケジュール」シートの状態とガントのバーを更新する
``UpdateWBS``) を、書き出したブックに入れて、マクロ有効の形式 (.xlsm) にする。

マクロは日本語のシート名 (「スケジュール」「設定」) を前提にしているので、組み込むのは
日本語で書き出すときだけ。環境変数 ``WBSGEN_MACRO=off`` で、組み込みをやめられる。
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from pathlib import Path
from typing import Optional

MACRO_FILE = Path(__file__).parent / "macros" / "vbaProject.bin"
ENV_SWITCH = "WBSGEN_MACRO"

VBA_PART = "xl/vbaProject.bin"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
CT_XLSM = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
CT_VBA = "application/vnd.ms-office.vbaProject"
REL_VBA = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"


def bundled(language: str = "ja") -> Optional[bytes]:
    """組み込むマクロの中身。組み込まない (言語が違う・無効・ファイルが無い) なら ``None``。"""
    if language != "ja":
        return None
    if os.environ.get(ENV_SWITCH, "").strip().lower() in ("off", "0", "no", "false"):
        return None
    try:
        return MACRO_FILE.read_bytes()
    except OSError:
        return None


def embed(path, vba: bytes) -> None:
    """``path`` の .xlsx にマクロを入れて、マクロ有効のブックにする (すでに入っていれば何もしない)。"""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}
    if VBA_PART in contents:
        return

    contents[VBA_PART] = vba
    contents["[Content_Types].xml"] = _content_types(contents["[Content_Types].xml"])
    contents["xl/_rels/workbook.xml.rels"] = _relationship(
        contents["xl/_rels/workbook.xml.rels"])
    contents["xl/workbook.xml"] = _code_name(contents["xl/workbook.xml"])

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in contents.items():
            zout.writestr(name, data)
    shutil.move(str(tmp), str(path))


def _content_types(data: bytes) -> bytes:
    xml = data.decode("utf-8").replace(CT_XLSX, CT_XLSM)
    if 'Extension="bin"' not in xml:
        xml = re.sub(r"(<Types[^>]*>)",
                     rf'\1<Default Extension="bin" ContentType="{CT_VBA}"/>', xml, count=1)
    return xml.encode("utf-8")


def _relationship(data: bytes) -> bytes:
    xml = data.decode("utf-8")
    used = {int(n) for n in re.findall(r'Id="rId(\d+)"', xml)}
    number = 1
    while number in used:
        number += 1
    entry = f'<Relationship Id="rId{number}" Type="{REL_VBA}" Target="vbaProject.bin"/>'
    return xml.replace("</Relationships>", entry + "</Relationships>").encode("utf-8")


def _code_name(data: bytes) -> bytes:
    """ブックのコード名 (``ThisWorkbook``) を付ける。マクロの部品がこの名前で結びつく。"""
    xml = data.decode("utf-8")
    if "codeName=" in xml.split("<sheets", 1)[0]:
        return data
    if re.search(r"<workbookPr\b", xml):
        xml = re.sub(r"<workbookPr\b", '<workbookPr codeName="ThisWorkbook"', xml, count=1)
    else:
        xml = re.sub(r"(<workbook\b[^>]*>)", r'\1<workbookPr codeName="ThisWorkbook"/>', xml,
                     count=1)
    return xml.encode("utf-8")
