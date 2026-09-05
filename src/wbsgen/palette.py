"""Excel の色を画面に出せる色に直す。

セルの文字色は RGB だけでなく、**色番号 (indexed)** や **テーマ色 (theme)**
でも指定される。Excel の色の選択で「テーマの色」の段を選ぶとテーマ色に、
古いファイルでは色番号になる。画面に出すにはどれも ``RRGGBB`` に直す必要が
ある。

書き出しでは openpyxl の色をそのまま書き戻すので、この変換は画面用。
"""

from __future__ import annotations

import colorsys
from typing import List, Optional
from xml.etree import ElementTree

from openpyxl.styles.colors import COLOR_INDEX

NS_DRAWING = "http://schemas.openxmlformats.org/drawingml/2006/main"

#: ``theme="N"`` が指す並び (clrMap の順)。theme1.xml に並んでいる順
#: (dk1, lt1, dk2, lt2, ...) とは、明暗の 2 組が入れ替わっている。
THEME_ORDER = (1, 0, 3, 2, 4, 5, 6, 7, 8, 9, 10, 11)

#: 色が決まらなかったときに返すもの (= Excel の「自動」)
AUTOMATIC = None


def theme_colors(book) -> List[str]:
    """ブックのテーマ色を ``theme="N"`` の順に並べて返す。

    テーマが読めなければ空を返す (テーマ色は画面では既定色になる)。
    """
    raw = getattr(book, "loaded_theme", None)
    if not raw:
        return []
    try:
        root = ElementTree.fromstring(raw)
    except Exception:  # noqa: BLE001 - 読めなければ諦める
        return []
    scheme = root.find(f".//{{{NS_DRAWING}}}clrScheme")
    if scheme is None:
        return []

    listed = []
    for slot in scheme:
        srgb = slot.find(f"{{{NS_DRAWING}}}srgbClr")
        system = slot.find(f"{{{NS_DRAWING}}}sysClr")
        value = (srgb.get("val") if srgb is not None
                 else system.get("lastClr") if system is not None else None)
        listed.append((value or "000000").upper()[-6:])
    return [listed[i] if i < len(listed) else "000000" for i in THEME_ORDER]


def to_rgb(color, theme: List[str] = ()) -> Optional[str]:
    """openpyxl の色を ``RRGGBB`` に直す。決まらなければ ``None``。"""
    if color is None:
        return AUTOMATIC
    base = _base(color, theme)
    if base is None:
        return AUTOMATIC
    tint = getattr(color, "tint", 0.0) or 0.0
    return _tinted(base, tint) if tint else base


def _base(color, theme: List[str]) -> Optional[str]:
    """濃淡 (tint) を掛ける前の色。"""
    kind = getattr(color, "type", None)
    if kind == "rgb":
        rgb = color.rgb
        return rgb.upper()[-6:] if isinstance(rgb, str) and len(rgb) >= 6 else None
    if kind == "indexed":
        index = color.indexed
        # 64 番以降は「自動」(前景色・背景色) なので色として扱わない
        if isinstance(index, int) and 0 <= index < len(COLOR_INDEX):
            return COLOR_INDEX[index].upper()[-6:]
        return None
    if kind == "theme":
        index = color.theme
        if isinstance(index, int) and 0 <= index < len(theme):
            return theme[index]
        return None
    return None


def _tinted(rgb: str, tint: float) -> str:
    """テーマ色の濃淡を掛ける (OOXML の決め方どおり、明度を動かす)。"""
    red, green, blue = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    hue, lum, sat = colorsys.rgb_to_hls(red, green, blue)
    lum = lum * (1 + tint) if tint < 0 else lum * (1 - tint) + tint
    red, green, blue = colorsys.hls_to_rgb(hue, min(max(lum, 0.0), 1.0), sat)
    return "%02X%02X%02X" % (round(red * 255), round(green * 255), round(blue * 255))
