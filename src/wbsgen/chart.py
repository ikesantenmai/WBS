"""ガントチャートの描画モデル。

読み込んだ WBS を、画面が描ける形 (行ごとのバーの位置) に変換する。
座標は「列単位の連続値」で返す。``x=2.5`` は 3 列目の中央を指し、
実際の px はブラウザ側が列幅を掛けて決める。
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Dict, List, Optional

from . import style
from .importer import ImportedWBS, Row
from .timeline import Timeline

#: 状態表示の色 (添付ファイルから抽出した配色)
STATUS_COLORS = {
    "完了": ("C0C0C0", "666666"),
    "実行中": ("FFFF99", "B36B00"),
    "残り": ("FFCC00", "800000"),
    "遅れ": ("FF99CC", "D00000"),
    "あと": ("CCFFFF", "3D8BC4"),
}


def build(imported: ImportedWBS, base_date: Optional[_dt.date] = None) -> Dict[str, Any]:
    """描画モデルを組み立てる。"""
    spec = imported.spec
    calendar = spec.calendar()
    timeline = Timeline(spec.start, spec.period_days, spec.unit, calendar)
    today = base_date or _dt.date.today()
    colors = _member_colors(imported)

    rows = [_row(row, timeline, calendar, colors) for row in imported.rows]
    return {
        "title": imported.title,
        "warnings": imported.warnings,
        "base_date": today.isoformat(),
        "now_x": timeline.position(today) if timeline.overlaps(today, today) else None,
        "timeline": _timeline(timeline),
        "rows": rows,
        "members": [{"name": name, "color": color} for name, color in colors.items()],
        "totals": _totals(imported.rows),
    }


# ----------------------------------------------------------------------
def _member_colors(imported: ImportedWBS) -> Dict[str, str]:
    """担当者に色を割り当てる。担当者一覧に載っている順を優先する。"""
    palette = style.MEMBER_PALETTE
    names: List[str] = list(imported.spec.members)
    for row in imported.rows:
        if row.member and row.member not in names:
            names.append(row.member)
    return {name: "#" + palette[i % len(palette)] for i, name in enumerate(names)}


def _timeline(timeline: Timeline) -> Dict[str, Any]:
    return {
        "unit": timeline.unit,
        "start": timeline.start.isoformat(),
        "end": timeline.end.isoformat(),
        "columns": [
            {
                "label": timeline.column_label(col),
                "weekday": timeline.weekday_label(col),
                "start": col.start.isoformat(),
                "rest": timeline.is_rest_column(col),
                "saturday": timeline.is_saturday_column(col),
            }
            for col in timeline.columns
        ],
        "bands": [
            {"start": index, "span": span, "label": label}
            for index, span, _day, label in timeline.header_top()
        ],
    }


# ----------------------------------------------------------------------
def _row(row: Row, timeline: Timeline, calendar, colors) -> Dict[str, Any]:
    plan_end = row.end
    if plan_end is None and row.start and row.days:
        # 終了日が書かれていなければ、日数 (稼働日) から補ってバーを描く
        plan_end = calendar.end_date(row.start, row.days)

    actual_end = row.actual_end
    if actual_end is None and row.actual_start and row.actual_days:
        actual_end = calendar.end_date(row.actual_start, row.actual_days)

    color = colors.get(row.member, "#" + style.MEMBER_PALETTE[0])
    background, foreground = _status_color(row.status)
    return {
        "row": row.row,
        "group": row.group,
        "subgroup": row.subgroup,
        "no": row.no,
        "name": row.name,
        "member": row.member,
        "color": color,
        "start": _iso(row.start),
        "days": row.days,
        "end": _iso(row.end or plan_end),
        "end_derived": row.end is None and plan_end is not None,
        "actual_start": _iso(row.actual_start),
        "actual_days": row.actual_days,
        "actual_end": _iso(row.actual_end),
        "delay": row.delay,
        "progress": row.progress,
        "effort": row.effort,
        "predecessor": row.predecessor,
        "status": row.status,
        "status_bg": background,
        "status_fg": foreground,
        "plan": _span(timeline, row.start, plan_end),
        "actual": _span(timeline, row.actual_start, actual_end),
    }


def _span(timeline: Timeline, start, end) -> Optional[Dict[str, float]]:
    """期間をチャート座標に直す。範囲外なら ``None``。"""
    if start is None or end is None or end < start:
        return None
    if not timeline.overlaps(start, end):
        return None
    return {
        "x1": timeline.position(timeline.clamp(start)),
        "x2": timeline.position(timeline.clamp(end), end_of_day=True),
    }


def _status_color(status: str):
    for key, value in STATUS_COLORS.items():
        if status.startswith(key):
            return "#" + value[0], "#" + value[1]
    return None, None


def _totals(rows: List[Row]) -> Dict[str, Any]:
    planned = [r for r in rows if r.has_plan or r.start]
    done = [r for r in rows if r.progress is not None and r.progress >= 1.0]
    running = [r for r in rows if r.has_actual
               and not (r.progress is not None and r.progress >= 1.0)]
    starts = [r.start for r in rows if r.start]
    ends = [r.end for r in rows if r.end]
    weight = sum(r.days or 1 for r in planned) or 1
    return {
        "rows": len(rows),
        "done": len(done),
        "running": len(running),
        "delayed": len([r for r in rows if r.delay]),
        "effort": round(sum(r.effort or 0 for r in rows), 1),
        "progress": round(
            sum((r.progress or 0) * (r.days or 1) for r in planned) / weight, 4),
        "first_day": _iso(min(starts)) if starts else None,
        "last_day": _iso(max(ends)) if ends else None,
    }


def _iso(value) -> Optional[str]:
    return value.isoformat() if value else None
