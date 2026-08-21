"""プロジェクトの辞書化。

:func:`wbsgen.loader.from_dict` の逆変換。Web UI との受け渡しと
YAML への書き戻しに使う。``from_dict(to_dict(p))`` は元と同じ内容になる。
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Dict

import yaml

from .model import KIND_TASK, Member, Project, Task


def to_dict(project: Project) -> Dict[str, Any]:
    """:class:`~wbsgen.model.Project` を JSON/YAML 化できる辞書にする。"""
    chart = project.chart
    show = chart.show
    return {
        "project": {
            "title": project.title,
            "updated_at": _dtstr(project.updated_at),
        },
        "chart": {
            "start": _datestr(chart.start),
            "period_days": chart.period_days,
            "unit": chart.unit,
            "base_date": _datestr(chart.base_date),
            "man_month_days": chart.man_month_days,
            "thresholds": {
                "exec_remain_days": chart.thresholds.exec_remain_days,
                "start_near_days": chart.thresholds.start_near_days,
            },
            "show": {k: getattr(show, k) for k in vars(show)},
        },
        "calendar": {
            "workdays": list(project.calendar.workdays),
            "japanese_holidays": project.calendar.japanese_holidays,
            "holidays": [_datestr(d) for d in project.calendar.holidays],
            "extra_workdays": [_datestr(d) for d in project.calendar.extra_workdays],
        },
        "members": [_member_dict(m) for m in project.members],
        "tasks": [_task_dict(t) for t in project.tasks],
        "blank_rows": project.blank_rows,
    }


def to_yaml(project: Project) -> str:
    """YAML テキストに書き出す (日本語はそのまま、キー順は定義順)。"""
    return yaml.safe_dump(
        to_dict(project),
        allow_unicode=True,
        sort_keys=False,
        default_flow_style=False,
        width=200,
    )


# ----------------------------------------------------------------------
def _member_dict(member: Member) -> Dict[str, Any]:
    data: Dict[str, Any] = {"name": member.name, "color": member.color}
    if member.workdays:
        data["workdays"] = list(member.workdays)
    if member.holidays:
        data["holidays"] = [_datestr(d) for d in member.holidays]
    if member.extra_workdays:
        data["extra_workdays"] = [_datestr(d) for d in member.extra_workdays]
    return data


def _task_dict(task: Task) -> Dict[str, Any]:
    """空の属性は落として、読みやすい辞書にする。"""
    data: Dict[str, Any] = {}
    for key in ("group", "subgroup", "no", "name"):
        value = getattr(task, key)
        if value:
            data[key] = value
    if task.kind != KIND_TASK:
        data["kind"] = task.kind
    for key in ("start", "end", "actual_start", "actual_end"):
        value = getattr(task, key)
        if value:
            data[key] = _datestr(value)
    for key in ("days", "actual_days", "progress", "effort"):
        value = getattr(task, key)
        if value is not None:
            data[key] = value
    for key in ("predecessor", "member", "comment", "shape"):
        value = getattr(task, key)
        if value:
            data[key] = value
    return data


def _datestr(value):
    if isinstance(value, _dt.datetime):
        value = value.date()
    return value.isoformat() if isinstance(value, _dt.date) else None


def _dtstr(value):
    if isinstance(value, _dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M")
    return _datestr(value)
