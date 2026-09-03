"""保存済み xlsx への DrawingML 注入。

openpyxl が書き出した zip を読み直し、指定シートに図形パートを追加する。
"""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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


def sheet_part(path: Path, title: str) -> str:
    """``title`` のシートが入っているパートの名前を返す。

    openpyxl はシートを並び順に ``sheet1.xml`` から書き出すが、元の
    ファイルに継ぎ足した場合は番号が並び順と一致しない。保存した
    ファイルの ``workbook.xml`` から引き当てる。
    """
    with zipfile.ZipFile(path) as archive:
        book = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rel_id = None
    for sheet in book.iter(f"{{{NS_MAIN}}}sheet"):
        if sheet.get("name") == title:
            rel_id = sheet.get(f"{{{NS_REL}}}id")
            break
    if rel_id is None:
        raise KeyError(title)

    for relationship in rels:
        if relationship.get("Id") == rel_id:
            # Target は "/xl/worksheets/sheet1.xml" とも
            # "worksheets/sheet1.xml" とも書かれる (書き手による)
            target = relationship.get("Target", "").lstrip("/")
            return target if target.startswith("xl/") else f"xl/{target}"
    raise KeyError(rel_id)


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
