"""DrawingML によるガントチャート図形の生成。

openpyxl はオートシェイプを扱えないため、xlsx (zip) に
``xl/drawings/drawingN.xml`` を直接組み立てて注入する。
元の Excel ツールがバーを図形として描いているのと同じ方式なので、
週表示・月表示でもセルの途中にバーの端を合わせられる。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Optional, Tuple
from xml.sax.saxutils import escape

EMU_PER_PX = 9525
EMU_PER_POINT = 12700

NS = (
    'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)


@dataclass
class Anchor:
    """セル基準のアンカー座標。"""

    col: int          # 0 始まりの列番号
    col_off: int      # 列内オフセット (EMU)
    row: int          # 0 始まりの行番号
    row_off: int      # 行内オフセット (EMU)

    def xml(self, tag: str) -> str:
        return (
            f"<xdr:{tag}><xdr:col>{self.col}</xdr:col>"
            f"<xdr:colOff>{int(self.col_off)}</xdr:colOff>"
            f"<xdr:row>{self.row}</xdr:row>"
            f"<xdr:rowOff>{int(self.row_off)}</xdr:rowOff></xdr:{tag}>"
        )


@dataclass
class Shape:
    """1 個の図形。"""

    frm: Anchor
    to: Anchor
    preset: str = "rect"
    fill: Optional[str] = None            # 単色 (RRGGBB)
    gradient: Optional[Tuple[str, str]] = None  # (濃い色, 淡い色)
    line_color: Optional[str] = None
    line_width_px: float = 0.75
    dash: Optional[str] = None            # "dash", "sysDot" など
    text: str = ""
    text_color: str = "000000"
    text_size: int = 10
    text_bold: bool = False
    text_align: str = "l"
    arrow_head: bool = False
    alpha: Optional[int] = None           # 透過率 (0-100000 の不透明度)
    adjust: Optional[float] = None        # プリセット図形の調整値
    flip_h: bool = False
    name: str = "shape"
    #: 絶対座標 (EMU)。:meth:`Drawing.finalize` が設定する。
    off: Tuple[int, int] = (0, 0)
    ext: Tuple[int, int] = (0, 0)


class Drawing:
    """1 シート分の図形コレクション。"""

    def __init__(self):
        self.shapes: List[Shape] = []

    def add(self, shape: Shape) -> Shape:
        self.shapes.append(shape)
        return shape

    def __len__(self) -> int:
        return len(self.shapes)

    def finalize(self, geometry: "Geometry") -> "Drawing":
        """各図形にシート上の絶対座標 (EMU) を書き込む。

        ``twoCellAnchor`` を使う場合でも Excel は ``a:xfrm`` の値を書き出すため、
        アンカーから実寸を計算して同じ形に揃えておく。
        """
        for shape in self.shapes:
            x1, y1 = geometry.absolute(shape.frm)
            x2, y2 = geometry.absolute(shape.to)
            shape.off = (x1, y1)
            shape.ext = (x2 - x1, y2 - y1)
        return self

    # ------------------------------------------------------------------
    def to_xml(self) -> str:
        parts = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr {NS}>']
        for i, shape in enumerate(self.shapes, start=2):
            parts.append(self._anchor_xml(shape, i))
        parts.append("</xdr:wsDr>")
        return "".join(parts)

    # ------------------------------------------------------------------
    def _anchor_xml(self, s: Shape, sid: int) -> str:
        body = [
            '<xdr:twoCellAnchor editAs="oneCell">',
            s.frm.xml("from"),
            s.to.xml("to"),
            "<xdr:sp macro=\"\" textlink=\"\">",
            "<xdr:nvSpPr>",
            f'<xdr:cNvPr id="{sid}" name="{escape(s.name)} {sid}"/>',
            "<xdr:cNvSpPr/>",
            "</xdr:nvSpPr>",
            "<xdr:spPr>",
            self._xfrm_xml(s),
            self._geom_xml(s),
            self._fill_xml(s),
            self._line_xml(s),
            "</xdr:spPr>",
            self._style_xml(),
            self._text_xml(s),
            "</xdr:sp>",
            "<xdr:clientData/>",
            "</xdr:twoCellAnchor>",
        ]
        return "".join(body)

    def _xfrm_xml(self, s: Shape) -> str:
        flip = ' flipH="1"' if s.flip_h else ""
        return (
            f"<a:xfrm{flip}>"
            f'<a:off x="{s.off[0]}" y="{s.off[1]}"/>'
            f'<a:ext cx="{max(s.ext[0], 0)}" cy="{max(s.ext[1], 0)}"/>'
            "</a:xfrm>"
        )

    def _geom_xml(self, s: Shape) -> str:
        if s.adjust is None:
            return f'<a:prstGeom prst="{s.preset}"><a:avLst/></a:prstGeom>'
        value = int(round(s.adjust * 100000))
        return (
            f'<a:prstGeom prst="{s.preset}"><a:avLst>'
            f'<a:gd name="adj" fmla="val {value}"/>'
            f"</a:avLst></a:prstGeom>"
        )

    def _fill_xml(self, s: Shape) -> str:
        alpha = f'<a:alpha val="{s.alpha}"/>' if s.alpha is not None else ""
        if s.gradient:
            dark, light = s.gradient
            return (
                "<a:gradFill flip=\"none\" rotWithShape=\"1\"><a:gsLst>"
                f'<a:gs pos="0"><a:srgbClr val="{light}">{alpha}</a:srgbClr></a:gs>'
                f'<a:gs pos="100000"><a:srgbClr val="{dark}">{alpha}</a:srgbClr></a:gs>'
                "</a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            )
        if s.fill:
            return f'<a:solidFill><a:srgbClr val="{s.fill}">{alpha}</a:srgbClr></a:solidFill>'
        return "<a:noFill/>"

    def _line_xml(self, s: Shape) -> str:
        if not s.line_color:
            return "<a:ln><a:noFill/></a:ln>"
        width = int(round(s.line_width_px * EMU_PER_PX))
        dash = f'<a:prstDash val="{s.dash}"/>' if s.dash else ""
        head = '<a:tailEnd type="triangle" w="sm" len="sm"/>' if s.arrow_head else ""
        return (
            f'<a:ln w="{width}" cap="flat"><a:solidFill>'
            f'<a:srgbClr val="{s.line_color}"/></a:solidFill>{dash}{head}</a:ln>'
        )

    def _style_xml(self) -> str:
        return (
            "<xdr:style>"
            '<a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef>'
            '<a:fillRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef>'
            '<a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef>'
            '<a:fontRef idx="none"/>'
            "</xdr:style>"
        )

    def _text_xml(self, s: Shape) -> str:
        if not s.text:
            return "<xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></xdr:txBody>"
        bold = "1" if s.text_bold else "0"
        size = s.text_size * 100
        return (
            '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="none" '
            'lIns="18000" tIns="0" rIns="18000" bIns="0" rtlCol="0" anchor="ctr"/>'
            "<a:lstStyle/>"
            f'<a:p><a:pPr algn="{s.text_align}"/>'
            f'<a:r><a:rPr lang="ja-JP" sz="{size}" b="{bold}">'
            f'<a:solidFill><a:srgbClr val="{s.text_color}"/></a:solidFill>'
            '<a:latin typeface="ＭＳ Ｐゴシック"/><a:ea typeface="ＭＳ Ｐゴシック"/>'
            f"</a:rPr><a:t>{escape(s.text)}</a:t></a:r></a:p></xdr:txBody>"
        )


# ----------------------------------------------------------------------
class Geometry:
    """行高・列幅から EMU 座標を求めるヘルパ。"""

    def __init__(self, col_widths_px, row_heights_pt):
        """
        Args:
            col_widths_px: 列番号 (0 始まり) -> 幅 px の写像。
            row_heights_pt: 行番号 (0 始まり) -> 高さ pt の写像。
        """
        self._cols = col_widths_px
        self._rows = row_heights_pt
        self._col_cache = {}
        self._row_cache = {}

    def absolute(self, anchor: Anchor) -> Tuple[int, int]:
        """アンカーをシート左上からの絶対 EMU 座標に変換する。"""
        return (
            self._col_prefix(anchor.col) + int(anchor.col_off),
            self._row_prefix(anchor.row) + int(anchor.row_off),
        )

    def _col_prefix(self, col: int) -> int:
        if col not in self._col_cache:
            total = self._col_prefix(col - 1) + self.col_width_emu(col - 1) if col else 0
            self._col_cache[col] = total
        return self._col_cache[col]

    def _row_prefix(self, row: int) -> int:
        if row not in self._row_cache:
            total = self._row_prefix(row - 1) + self.row_height_emu(row - 1) if row else 0
            self._row_cache[row] = total
        return self._row_cache[row]

    def col_width_emu(self, col: int) -> int:
        return int(round(self._cols(col) * EMU_PER_PX))

    def row_height_emu(self, row: int) -> int:
        return int(round(self._rows(row) * EMU_PER_POINT))

    # ------------------------------------------------------------------
    def anchor(self, col_pos: float, row: int, row_frac: float = 0.0,
               first_col: int = 0) -> Anchor:
        """「列単位の連続座標」と行番号からアンカーを作る。

        Args:
            col_pos: チャート左端を 0 とした列単位の座標 (小数可)。
            row: 0 始まりの行番号。
            row_frac: 行内の縦位置 (0.0=上端, 1.0=下端)。
            first_col: チャート先頭列の 0 始まり列番号。
        """
        col_index = int(col_pos)
        frac = col_pos - col_index
        abs_col = first_col + col_index
        col_off = int(round(frac * self.col_width_emu(abs_col)))
        row_off = int(round(row_frac * self.row_height_emu(row)))
        return Anchor(abs_col, col_off, row, row_off)
