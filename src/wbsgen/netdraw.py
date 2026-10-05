"""ネットワーク図 (プレジデンス図) の図形の組み立て。

画面と同じ配置 (:mod:`wbsgen.network` が決めた段とレーン) を、Excel の
図形に起こす。箱は四角形、前後関係は矢印つきの折れ線にして、上の段に
ES / 日数 / EF、下の段に LS / 余裕 / LF を並べる。

シートの列幅と行の高さはすべて同じ (:data:`CELL_PX`) にしてあるので、
画面の px 座標をそのままアンカーに直せる。
"""

from __future__ import annotations

from typing import Any, Dict, Tuple

from .drawing import EMU_PER_PX, Anchor, Drawing, Geometry, Shape

CELL_PX = 20
CELL_PT = CELL_PX * 0.75

BOX_W = 208
BOX_H = 84
GAP_X = 64
GAP_Y = 16
PAD = 20

INK = "000000"
MUTED = "595959"
LINE = "7F7F7F"
LATE = "C00000"


def geometry() -> Geometry:
    return Geometry(lambda _col: CELL_PX, lambda _row: CELL_PT)


def size(network: Dict[str, Any]) -> Tuple[int, int]:
    """図の幅と高さ (px)。"""
    nodes = network["nodes"]
    if not nodes:
        return 0, 0
    levels = max(node["level"] for node in nodes) + 1
    lanes = max(node["lane"] for node in nodes) + 1
    return (PAD * 2 + levels * BOX_W + (levels - 1) * GAP_X,
            PAD * 2 + lanes * BOX_H + (lanes - 1) * GAP_Y)


def _anchor(x: float, y: float) -> Anchor:
    col, col_off = divmod(int(round(x)), CELL_PX)
    row, row_off = divmod(int(round(y)), CELL_PX)
    return Anchor(col, col_off * EMU_PER_PX, row, row_off * EMU_PER_PX)


def _md(iso: str) -> str:
    _year, month, day = iso.split("-")
    return f"{int(month)}/{int(day)}"


def build(network: Dict[str, Any], top_px: int = 0) -> Drawing:
    """プレジデンス図の図形。``top_px`` は図の上端の位置 (見出しのぶんだけ下げる)。"""
    drawing = Drawing()
    by_row = {node["row"]: node for node in network["nodes"]}

    def pos(node):
        return (PAD + node["level"] * (BOX_W + GAP_X),
                top_px + PAD + node["lane"] * (BOX_H + GAP_Y))

    def add(x1, y1, x2, y2, **kwargs):
        drawing.add(Shape(frm=_anchor(x1, y1), to=_anchor(x2, y2), **kwargs))

    # 矢印を先に描き、箱をその上に重ねる
    for edge in network["edges"]:
        source, target = by_row.get(edge["from"]), by_row.get(edge["to"])
        if not source or not target:
            continue
        sx, sy = pos(source)
        tx, ty = pos(target)
        x1, y1 = sx + BOX_W, sy + BOX_H / 2
        x2, y2 = tx, ty + BOX_H / 2
        late, critical = edge.get("late"), edge.get("critical")
        add(x1, min(y1, y2), x2, max(y1, y2),
            preset="bentConnector3", flip_v=y2 < y1, arrow_head=True,
            line_color=LATE if late else (INK if critical else LINE),
            line_width_px=2 if (late or critical) else 1.25,
            name=f"edge-{edge['from']}-{edge['to']}")

    for node in network["nodes"]:
        x, y = pos(node)
        fill = (node.get("status_bg") or "#FFFFFF").lstrip("#")
        ink = (node.get("status_fg") or "#000000").lstrip("#")
        label = node["no"]
        add(x, y, x + BOX_W, y + BOX_H, preset="rect", fill=fill,
            line_color=INK if node.get("critical") else "A6A6A6",
            line_width_px=2.25 if node.get("critical") else 0.75,
            name=f"node-{label}")

        inner = BOX_W - 12
        cell = (inner - 8) / 3

        def row_cells(top, values):
            for i, value in enumerate(values):
                left = x + 6 + i * (cell + 4)
                add(left, top, left + cell, top + 18, preset="rect",
                    fill="FFFFFF", line_color="808080", text=str(value),
                    text_size=8, text_bold=i != 1, text_align="ctr",
                    text_color=INK, name=f"time-{label}")

        row_cells(y + 4, [node["es"], node["duration"], node["ef"]])
        add(x + 6, y + 26, x + BOX_W - 6, y + 42, preset="rect",
            text=f"{label} {node['name']}".strip(), text_size=9, text_bold=True,
            text_color=ink, name=f"name-{label}")
        period = (f"{_md(node['start'])} - {_md(node['end'])}"
                  if node.get("start") and node.get("end") else "")
        add(x + 6, y + 44, x + 6 + inner * 0.6, y + 60, preset="rect",
            text=period, text_size=8, text_color=MUTED, name=f"period-{label}")
        add(x + 6 + inner * 0.6, y + 44, x + BOX_W - 6, y + 60, preset="rect",
            text=node.get("member") or "", text_size=8, text_align="r",
            text_color=MUTED, name=f"member-{label}")
        row_cells(y + 62, [node["ls"], node["float"], node["lf"]])
    return drawing
