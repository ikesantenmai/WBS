"""ガントチャートの描画モデル。

読み込んだ WBS を、画面が描ける形 (行ごとのバーの位置) に変換する。
座標は「列単位の連続値」で返す。``x=2.5`` は 3 列目の中央を指し、
実際の px はブラウザ側が列幅を掛けて決める。
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Dict, List, Optional

from . import style
from .i18n import labels as get_labels, status_kind
from .importer import ImportedWBS, Row, resolve
from .timeline import Timeline
from .workload import build as build_workload

#: 画面での状態表示の色 (添付ファイルの配色を、画面で読みやすいように調整)
STATUS_COLORS = {
    "done": ("C0C0C0", "666666"),
    "running": ("FFFF99", "B36B00"),
    "remaining": ("FFCC00", "800000"),
    "delayed": ("FF99CC", "D00000"),
    "upcoming": ("CCFFFF", "3D8BC4"),
}


def build(imported: ImportedWBS, base_date: Optional[_dt.date] = None) -> Dict[str, Any]:
    """描画モデルを組み立てる。"""
    spec = imported.spec
    calendar = spec.calendar()
    timeline = Timeline(spec.start, spec.period_days, spec.unit, calendar, spec.language)
    today = base_date or _dt.date.today()
    # 状態と遅れは基準日で決まるので、この日付で数え直す
    resolve(imported.rows, calendar, today, spec.language)
    colors = _member_colors(imported)

    rows = [_row(row, timeline, colors) for row in imported.rows]
    return {
        "title": imported.title,
        "language": spec.language,
        "warnings": imported.warnings,
        "base_date": today.isoformat(),
        "now_x": timeline.position(today) if timeline.overlaps(today, today) else None,
        "timeline": _timeline(timeline),
        "rows": rows,
        "members": [{"name": name, "color": color} for name, color in colors.items()],
        "totals": _totals(rows),
        "workload": _workload(imported, calendar, spec.language),
    }


def _workload(imported: ImportedWBS, calendar, language: str) -> List[Dict[str, Any]]:
    """要員稼働チェックを、画面に出せる形にする (Excel のシートと同じ中身)。"""
    text = get_labels(language)
    dated = _repeats(months := build_workload(imported.rows, calendar))
    template = text.load_sheet_dated if dated else text.load_sheet
    return [
        {
            "label": template.format(year=month.year, month=month.month),
            "title": text.load_title.format(
                start=_md(month.start), end=_md(month.end, month.start)),
            "days": [
                {
                    "date": day.isoformat(),
                    "label": f"{day.month}/{day.day}",
                    "weekday": text.weekdays[day.weekday()],
                    "working": working,
                }
                for day, working in zip(month.days, month.workdays)
            ],
            "workdays": month.workday_count,
            "members": [
                {
                    "name": load.name,
                    "counts": load.counts,
                    "busy": load.busy_days,
                    "free": len(load.free_days),
                    "free_days": ", ".join(
                        text.load_free_item.format(
                            month=day.month, day=day.day,
                            weekday=text.weekdays[day.weekday()])
                        for day in load.free_days),
                }
                for load in month.members
            ],
        }
        for month in months
    ]


def _repeats(months) -> bool:
    """同じ月が 2 度出てくるか (名前に年を入れるかの判断)。"""
    seen = [month.month for month in months]
    return len(seen) != len(set(seen))


def _md(day: _dt.date, since: _dt.date = None) -> str:
    """見出しに出す日付。同じ年なら年を省く。"""
    if since is not None and since.year == day.year:
        return f"{day.month}/{day.day}"
    return f"{day.year}/{day.month}/{day.day}"


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
def _row(row: Row, timeline: Timeline, colors) -> Dict[str, Any]:
    """日数・終了日・進捗は :func:`wbsgen.importer.resolve` が解決済み。"""
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
        "end": _iso(row.end),
        "actual_start": _iso(row.actual_start),
        "actual_days": row.actual_days,
        "actual_end": _iso(row.actual_end),
        # 記入内容から導き出した項目 (画面で薄く見せる)
        "derived": sorted(row.derived),
        # セルに指定されていた文字色 (画面でもそのまま見せる)。
        # RGB のほか、色番号やテーマ色も画面用の色に直してある。
        "colors": {key: "#" + rgb for key, rgb in row.ink.items()},
        # セルの背景も、書かれていたものをそのまま見せる
        "fills": {key: "#" + rgb for key, rgb in row.paper.items()},
        "delay": row.delay,
        "progress": row.progress,
        "effort": row.effort,
        "predecessor": row.predecessor,
        "status": row.status,
        "status_bg": background,
        "status_fg": foreground,
        "plan": _span(timeline, row.start, row.end),
        "actual": _span(timeline, row.actual_start, row.actual_end),
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
    """状態の色。書かれた言葉から種類を判定する (日英どちらでも効く)。"""
    kind = status_kind(status)
    if kind is None:
        return None, None
    background, foreground = STATUS_COLORS[kind]
    return "#" + background, "#" + foreground


def _totals(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """描画済みの行から集計する。

    終了日は補ったものを使う (画面に描いたバーと期間を一致させるため)。
    """
    planned = [r for r in rows if r["start"]]
    done = [r for r in rows if r["progress"] is not None and r["progress"] >= 1.0]
    running = [r for r in rows if r["actual_start"]
               and not (r["progress"] is not None and r["progress"] >= 1.0)]
    starts = [r["start"] for r in rows if r["start"]]
    ends = [r["end"] for r in rows if r["end"]]
    weight = sum(r["days"] or 1 for r in planned) or 1
    return {
        "rows": len(rows),
        "done": len(done),
        "running": len(running),
        "delayed": len([r for r in rows if r["delay"]]),
        "effort": round(sum(r["effort"] or 0 for r in rows), 1),
        "progress": round(
            sum((r["progress"] or 0) * (r["days"] or 1) for r in planned) / weight, 4),
        "first_day": min(starts) if starts else None,
        "last_day": max(ends) if ends else None,
    }


def _iso(value) -> Optional[str]:
    return value.isoformat() if value else None
