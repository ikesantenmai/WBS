"""ブラウザ描画用のチャートモデル。

Excel 出力と同じエンジン (:mod:`wbsgen.scheduling` と
:class:`~wbsgen.render.timeline.Timeline`) から座標を作るので、
画面のプレビューと生成される Excel が一致する。

座標は「列単位の連続値」で返す。``x=2.5`` は 3 列目の中央を指す。
実際の px はブラウザ側が列幅を掛けて決める。
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Dict, List, Optional

from . import style
from .model import KIND_MILESTONE, KIND_SUMMARY, MILESTONE_SHAPES, Project
from .render.timeline import Timeline
from .scheduling import inazuma_points, resolve_project


def build_preview(project: Project) -> Dict[str, Any]:
    """解決済みのチャートモデルを返す。``project`` は破壊的に更新される。"""
    calendar = resolve_project(project)
    chart = project.chart
    timeline = Timeline(chart.start, chart.period_days, chart.unit, calendar)
    base_date = chart.base_date or _dt.date.today()
    colors = _member_colors(project)

    rows = []
    for index, task in enumerate(project.tasks):
        task.row = index + 1
        rows.append(_row(task, timeline, colors, project))

    return {
        "title": project.title,
        "base_date": base_date.isoformat(),
        "timeline": _timeline(timeline),
        "rows": rows,
        "links": _links(project, rows),
        "now_x": timeline.position(base_date) if timeline.in_range(base_date, base_date) else None,
        "inazuma": _inazuma(project, timeline, calendar, base_date),
        "totals": _totals(project),
        "show": {k: getattr(chart.show, k) for k in vars(chart.show)},
    }


# ----------------------------------------------------------------------
def _member_colors(project: Project) -> Dict[str, str]:
    palette = list(style.MEMBER_PALETTE)
    return {
        m.name: "#" + (m.color or palette[i % len(palette)]).lstrip("#").upper()
        for i, m in enumerate(project.members)
    }


def _color_for(task, colors: Dict[str, str]) -> str:
    color = colors.get(task.member)
    if color:
        return color
    if task.member:
        index = sum(ord(c) for c in task.member) % len(style.MEMBER_PALETTE)
        return "#" + style.MEMBER_PALETTE[index]
    return "#" + style.MEMBER_PALETTE[0]


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
            {"start": start, "span": span, "label": label}
            for start, span, label in timeline.band_labels()
        ],
    }


def _span(timeline: Timeline, start, end) -> Optional[Dict[str, float]]:
    if start is None or end is None or not timeline.in_range(start, end):
        return None
    return {
        "x1": timeline.position(timeline.clamp(start)),
        "x2": timeline.position(timeline.clamp(end), end_of_day=True),
    }


def _row(task, timeline: Timeline, colors, project) -> Dict[str, Any]:
    color = _color_for(task, colors)
    actual_end = task.actual_end
    if task.actual_start and actual_end is None:
        base = project.chart.base_date or _dt.date.today()
        actual_end = max(task.actual_start, base)

    data = {
        "row": task.row,
        "group": task.group,
        "subgroup": task.subgroup,
        "no": task.no,
        "name": task.name,
        "kind": task.kind,
        "member": task.member,
        "color": color,
        "start": _iso(task.start),
        "days": task.days,
        "end": _iso(task.end),
        "actual_start": _iso(task.actual_start),
        "actual_days": task.actual_days,
        "actual_end": _iso(task.actual_end),
        "progress": task.progress,
        "effort": task.effort,
        "predecessor": task.predecessor,
        "delay": task.delay,
        "status": task.status,
        "comment": task.comment,
        "plan": None,
        "actual": None,
        "milestone": None,
    }
    if task.kind == KIND_MILESTONE:
        day = task.start or task.actual_start
        if day and timeline.in_range(day, day):
            data["milestone"] = {
                "x": timeline.position(timeline.clamp(day)),
                "shape": MILESTONE_SHAPES.get(task.shape or "diamond", "diamond"),
            }
    else:
        data["plan"] = _span(timeline, task.start, task.end)
        if project.chart.show.results and task.actual_start:
            data["actual"] = _span(timeline, task.actual_start, actual_end)
        if task.kind == KIND_SUMMARY and data["plan"]:
            data["plan"]["summary"] = True
    return data


def _links(project: Project, rows) -> List[Dict[str, Any]]:
    if not project.chart.show.predecessor_line:
        return []
    by_no = {str(t.no): t for t in project.tasks if t.no}
    spans = {r["row"]: r for r in rows}
    links = []
    for task in project.tasks:
        if not task.predecessor:
            continue
        target = spans.get(task.row)
        if not target:
            continue
        start_x = _left_edge(target)
        if start_x is None:
            continue
        for key in str(task.predecessor).replace("、", ",").split(","):
            pred = by_no.get(key.strip())
            if pred is None:
                continue
            source = spans.get(pred.row)
            end_x = _right_edge(source) if source else None
            if end_x is None:
                continue
            links.append({
                "from_row": pred.row, "to_row": task.row,
                "x1": end_x, "x2": start_x,
            })
    return links


def _left_edge(row) -> Optional[float]:
    if row["plan"]:
        return row["plan"]["x1"]
    if row["milestone"]:
        return row["milestone"]["x"]
    return None


def _right_edge(row) -> Optional[float]:
    if row["plan"]:
        return row["plan"]["x2"]
    if row["milestone"]:
        return row["milestone"]["x"]
    return None


def _inazuma(project: Project, timeline: Timeline, calendar, base_date):
    if not project.chart.show.inazuma_line:
        return []
    points = []
    for task, achieved in inazuma_points(project, calendar, base_date):
        if not timeline.in_range(achieved, achieved):
            continue
        points.append({
            "row": task.row,
            "x": timeline.position(timeline.clamp(achieved), end_of_day=task.is_complete),
        })
    return points


def _totals(project: Project) -> Dict[str, Any]:
    tasks = [t for t in project.tasks if t.kind != KIND_SUMMARY]
    done = [t for t in tasks if t.is_complete]
    running = [t for t in tasks if t.actual_start and not t.is_complete]
    delayed = [t for t in tasks if t.delay]
    ends = [t.end for t in tasks if t.end]
    starts = [t.start for t in tasks if t.start]
    return {
        "tasks": len(tasks),
        "done": len(done),
        "running": len(running),
        "delayed": len(delayed),
        "effort": round(sum(t.effort or 0 for t in tasks), 1),
        "first_day": _iso(min(starts)) if starts else None,
        "last_day": _iso(max(ends)) if ends else None,
        "progress": round(
            sum((t.progress or 0) * (t.days or 1) for t in tasks)
            / max(sum(t.days or 1 for t in tasks), 1), 4),
    }


def _iso(value):
    return value.isoformat() if value else None
