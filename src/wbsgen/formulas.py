"""セルの数式を、書き出し先の位置に合わせて書き直す。

読み込んだ表と、書き出す表では、行や列の位置が違うことがある (空行を詰める、
列の並びがこのツールの形に変わる、など)。そのまま写すと、数式が別のセルを
指してしまうので、同じシートの参照だけを、新しい位置に直す。

直せない数式 (空行や見出しなど、書き出さないセルを指しているもの) は ``None``
にして、呼び出し側が数式のかわりに値を書けるようにする。
"""

from __future__ import annotations

import re
from typing import Dict, Optional

from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter

CELL = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")


def translate(formula: str, columns: Dict[int, int], rows: Dict[int, int],
              sheet_title: str) -> Optional[str]:
    """``formula`` の同じシートの参照を、``columns`` / ``rows`` (元 -> 先) で直す。

    他のシートへの参照はそのまま。直せない参照があれば ``None``。
    """
    try:
        tokens = Tokenizer(formula).items
    except Exception:  # noqa: BLE001 - 読めない数式は、直せないものとして扱う
        return None
    out = ["="]
    for token in tokens:
        text = token.value
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            text = _reference(text, columns, rows, sheet_title)
            if text is None:
                return None
        out.append(text)
    return "".join(out)


def moves(columns: Dict[int, int], rows: Dict[int, int]) -> bool:
    """列や行の位置が変わるものがあるか (変わらなければ数式はそのまま写せる)。"""
    return any(a != b for a, b in columns.items()) or any(a != b for a, b in rows.items())


def _reference(text: str, columns: Dict[int, int], rows: Dict[int, int],
               sheet_title: str) -> Optional[str]:
    sheet, bang, ref = text.rpartition("!")
    if bang and sheet.strip("'").replace("''", "'") != sheet_title:
        return text                                   # 別のシートはそのまま
    parts = ref.split(":")
    out = []
    for part in parts:
        found = CELL.match(part)
        if not found:
            # 名前 (定義済みの名前など) は、そのまま。範囲の一部が列だけ・行だけなら直せない
            if len(parts) == 1:
                return text
            return None
        d_col, letter, d_row, number = found.groups()
        col = columns.get(column_index_from_string(letter.upper()))
        row = rows.get(int(number))
        if col is None or row is None:
            return None
        out.append(f"{d_col}{get_column_letter(col)}{d_row}{row}")
    return (sheet + bang if bang else "") + ":".join(out)
