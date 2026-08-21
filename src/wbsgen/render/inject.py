"""保存済み xlsx への DrawingML 注入。

openpyxl が書き出した zip を読み直し、指定シートに図形パートを追加する。
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

CT_DRAWING = "application/vnd.openxmlformats-officedocument.drawing+xml"
REL_DRAWING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


def inject_drawing(path: Path, sheet_part: str, drawing_xml: str) -> None:
    """``sheet_part`` (例: ``xl/worksheets/sheet1.xml``) に図形を追加する。"""
    path = Path(path)
    tmp = path.with_suffix(path.suffix + ".tmp")

    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    index = _next_drawing_index(names)
    drawing_part = f"xl/drawings/drawing{index}.xml"
    rels_part = _rels_part(sheet_part)
    rel_id = _add_relationship(contents, rels_part, drawing_part)

    contents[drawing_part] = drawing_xml.encode("utf-8")
    contents["[Content_Types].xml"] = _add_content_type(
        contents["[Content_Types].xml"], drawing_part
    )
    contents[sheet_part] = _add_drawing_ref(contents[sheet_part], rel_id)

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in contents.items():
            zout.writestr(name, data)
    shutil.move(str(tmp), str(path))


# ----------------------------------------------------------------------
def _next_drawing_index(names) -> int:
    used = {
        int(m.group(1))
        for n in names
        for m in [re.fullmatch(r"xl/drawings/drawing(\d+)\.xml", n)]
        if m
    }
    i = 1
    while i in used:
        i += 1
    return i


def _rels_part(sheet_part: str) -> str:
    head, _, tail = sheet_part.rpartition("/")
    return f"{head}/_rels/{tail}.rels"


def _add_relationship(contents, rels_part: str, target_part: str) -> str:
    """シートの .rels に drawing への関連を追加し、その rId を返す。"""
    if rels_part in contents:
        xml = contents[rels_part].decode("utf-8")
    else:
        xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            "</Relationships>"
        )
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', xml)}
    rid = 1
    while rid in used:
        rid += 1
    rel_id = f"rId{rid}"
    # drawing のパスはシートから見た相対パス (../drawings/drawingN.xml)
    target = "../" + target_part[len("xl/"):]
    entry = f'<Relationship Id="{rel_id}" Type="{REL_DRAWING}" Target="{target}"/>'
    xml = xml.replace("</Relationships>", entry + "</Relationships>")
    contents[rels_part] = xml.encode("utf-8")
    return rel_id


def _add_content_type(data: bytes, part: str) -> bytes:
    xml = data.decode("utf-8")
    override = f'<Override PartName="/{part}" ContentType="{CT_DRAWING}"/>'
    if override in xml:
        return data
    return xml.replace("</Types>", override + "</Types>").encode("utf-8")


R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _add_drawing_ref(data: bytes, rel_id: str) -> bytes:
    """worksheet XML に ``<drawing r:id="..."/>`` を正しい位置で挿入する。

    ``<drawing>`` はスキーマ上 ``<legacyDrawing>``/``<picture>`` の直前、
    かつ ``<pageSetup>`` などの後ろに置く必要がある。
    openpyxl は関連を持たないシートに ``r`` 名前空間を宣言しないため、
    未宣言なら ``<worksheet>`` 要素に補う。
    """
    xml = data.decode("utf-8")
    if "<drawing " in xml:
        return data
    if 'xmlns:r=' not in xml.split(">", 2)[0] + ">":
        xml = re.sub(
            r"(<worksheet\b)",
            r'\1 xmlns:r="%s"' % R_NS,
            xml,
            count=1,
        )
    tag = f'<drawing r:id="{rel_id}"/>'
    for anchor in ("<legacyDrawing", "<picture", "<oleObjects", "<extLst"):
        pos = xml.find(anchor)
        if pos != -1:
            return (xml[:pos] + tag + xml[pos:]).encode("utf-8")
    return xml.replace("</worksheet>", tag + "</worksheet>").encode("utf-8")
