"""プロジェクト定義の読み込み。

YAML / JSON をそのまま、CSV はタスク表として読み込む。
"""

from __future__ import annotations

import csv
import datetime as _dt
import json
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml

from .model import (
    KIND_TASK,
    VALID_KINDS,
    CalendarConfig,
    ChartConfig,
    DisplayFlags,
    Member,
    Project,
    StatusThresholds,
    Task,
)


class ProjectError(ValueError):
    """プロジェクト定義の不備。"""


# ----------------------------------------------------------------------
def load(path) -> Project:
    """拡張子に応じて YAML / JSON / CSV を読み込む。"""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in (".yaml", ".yml"):
        return from_dict(yaml.safe_load(path.read_text(encoding="utf-8")) or {})
    if suffix == ".json":
        return from_dict(json.loads(path.read_text(encoding="utf-8")))
    if suffix in (".csv", ".tsv"):
        return from_csv(path)
    raise ProjectError(f"対応していない拡張子です: {path.suffix} (.yaml/.json/.csv)")


# ----------------------------------------------------------------------
def from_dict(data: Dict[str, Any]) -> Project:
    if not isinstance(data, dict):
        raise ProjectError("プロジェクト定義はマッピングである必要があります。")

    meta = data.get("project") or {}
    project = Project(
        title=meta.get("title", "プロジェクト スケジュール"),
        updated_at=_datetime(meta.get("updated_at")),
        chart=_chart(data.get("chart") or {}),
        calendar=_calendar(data.get("calendar") or {}),
        members=[_member(m) for m in (data.get("members") or [])],
        tasks=[_task(t, i) for i, t in enumerate(data.get("tasks") or [])],
        standard_process=[_task(t, i) for i, t in enumerate(data.get("standard_process") or [])],
    )
    if not project.tasks:
        raise ProjectError("tasks が空です。1 件以上のタスクを定義してください。")
    _fill_inherited(project.tasks)
    _fill_inherited(project.standard_process)
    return project


def _chart(data: Dict[str, Any]) -> ChartConfig:
    show = DisplayFlags()
    for key, value in (data.get("show") or {}).items():
        if not hasattr(show, key):
            raise ProjectError(f"chart.show に未知のキーがあります: {key}")
        setattr(show, key, bool(value))

    thresholds = StatusThresholds()
    for key, value in (data.get("thresholds") or {}).items():
        if not hasattr(thresholds, key):
            raise ProjectError(f"chart.thresholds に未知のキーがあります: {key}")
        setattr(thresholds, key, int(value))

    return ChartConfig(
        start=_date(data.get("start")),
        period_days=int(data.get("period_days", 360)),
        unit=str(data.get("unit", "week")),
        base_date=_date(data.get("base_date")),
        man_month_days=int(data.get("man_month_days", 20)),
        thresholds=thresholds,
        show=show,
    )


def _calendar(data: Dict[str, Any]) -> CalendarConfig:
    return CalendarConfig(
        workdays=list(data.get("workdays") or ["mon", "tue", "wed", "thu", "fri"]),
        japanese_holidays=bool(data.get("japanese_holidays", True)),
        holidays=[_date(d) for d in (data.get("holidays") or [])],
        extra_workdays=[_date(d) for d in (data.get("extra_workdays") or [])],
    )


def _member(data) -> Member:
    if isinstance(data, str):
        return Member(name=data)
    return Member(
        name=str(data.get("name", "")),
        color=str(data.get("color", "#4472C4")),
        workdays=list(data["workdays"]) if data.get("workdays") else None,
        holidays=[_date(d) for d in (data.get("holidays") or [])],
        extra_workdays=[_date(d) for d in (data.get("extra_workdays") or [])],
    )


#: 属性名の別名。``no`` は YAML 1.1 では真偽値 ``False`` に解決されるため、
#: 素の ``no:`` キーも項番として受け付ける。
KEY_ALIASES = {False: "no", "id": "no", True: "yes"}


def _normalize_keys(data: Dict[Any, Any]) -> Dict[str, Any]:
    return {KEY_ALIASES.get(k, k): v for k, v in data.items()}


def _task(data: Dict[str, Any], index: int) -> Task:
    if not isinstance(data, dict):
        raise ProjectError(f"tasks[{index}] はマッピングである必要があります。")
    data = _normalize_keys(data)
    kind = str(data.get("kind", KIND_TASK))
    if kind not in VALID_KINDS:
        raise ProjectError(f"tasks[{index}].kind が不正です: {kind} ({'/'.join(VALID_KINDS)})")
    name = data.get("name")
    if not name:
        raise ProjectError(f"tasks[{index}] に name がありません。")
    return Task(
        name=str(name),
        group=str(data.get("group", "") or ""),
        subgroup=str(data.get("subgroup", "") or ""),
        no=_no(data.get("no")),
        kind=kind,
        start=_date(data.get("start")),
        days=_int(data.get("days"), f"tasks[{index}].days"),
        end=_date(data.get("end")),
        actual_start=_date(data.get("actual_start")),
        actual_days=_int(data.get("actual_days"), f"tasks[{index}].actual_days"),
        actual_end=_date(data.get("actual_end")),
        progress=_progress(data.get("progress"), index),
        effort=_float(data.get("effort"), f"tasks[{index}].effort"),
        predecessor=_no(data.get("predecessor")),
        member=str(data.get("member", "") or ""),
        comment=str(data.get("comment", "") or ""),
        shape=str(data.get("shape", "") or ""),
    )


def _fill_inherited(tasks: List[Task]) -> None:
    """``group`` / ``subgroup`` が空の行は直前の行の値を引き継ぐ。"""
    group = subgroup = ""
    for task in tasks:
        if task.group:
            group = task.group
            if not task.subgroup:
                subgroup = ""
        else:
            task.group = group
        if task.subgroup:
            subgroup = task.subgroup
        else:
            task.subgroup = subgroup


# ----------------------------------------------------------------------
def from_csv(path: Path, encoding: str = "utf-8-sig") -> Project:
    """タスク表 CSV からプロジェクトを組み立てる (設定は既定値)。"""
    rows = read_task_csv(path, encoding=encoding)
    members = sorted({r["member"] for r in rows if r.get("member")})
    return from_dict({
        "project": {"title": path.stem},
        "members": [{"name": m} for m in members],
        "tasks": rows,
    })


#: CSV 列名 -> タスク属性 (日本語ヘッダにも対応)
CSV_FIELDS = {
    "group": "group", "大項目": "group",
    "subgroup": "subgroup", "中項目": "subgroup",
    "no": "no", "項番": "no",
    "name": "name", "項目": "name", "タスク": "name",
    "kind": "kind", "種別": "kind",
    "start": "start", "開始日": "start", "予定開始": "start",
    "days": "days", "日数": "days", "予定日数": "days",
    "end": "end", "終了日": "end", "予定終了": "end",
    "actual_start": "actual_start", "実績開始": "actual_start",
    "actual_days": "actual_days", "実績日数": "actual_days",
    "actual_end": "actual_end", "実績終了": "actual_end",
    "progress": "progress", "進捗": "progress",
    "effort": "effort", "工数": "effort",
    "predecessor": "predecessor", "先行": "predecessor",
    "member": "member", "担当": "member",
    "comment": "comment", "コメント": "comment",
    "shape": "shape", "記号": "shape",
}


def read_task_csv(path: Path, encoding: str = "utf-8-sig") -> List[Dict[str, Any]]:
    delimiter = "\t" if Path(path).suffix.lower() == ".tsv" else ","
    with open(path, newline="", encoding=encoding) as handle:
        reader = csv.DictReader(handle, delimiter=delimiter)
        if reader.fieldnames is None:
            raise ProjectError(f"{path} にヘッダ行がありません。")
        unknown = [f for f in reader.fieldnames if f and f.strip() not in CSV_FIELDS]
        if unknown:
            raise ProjectError(
                f"CSV に未知の列があります: {', '.join(unknown)}\n"
                f"使用できる列: {', '.join(sorted(set(CSV_FIELDS)))}"
            )
        rows = []
        for raw in reader:
            row = {}
            for key, value in raw.items():
                if key is None or value is None or value.strip() == "":
                    continue
                row[CSV_FIELDS[key.strip()]] = value.strip()
            if row.get("name"):
                rows.append(row)
    if not rows:
        raise ProjectError(f"{path} に有効なタスク行がありません。")
    return rows


# ----------------------------------------------------------------------
# 値の変換
# ----------------------------------------------------------------------
def _date(value) -> Optional[_dt.date]:
    if value in (None, ""):
        return None
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    text = str(value).strip().replace("/", "-")
    for fmt in ("%Y-%m-%d", "%Y-%m-%d %H:%M", "%y-%m-%d"):
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ProjectError(f"日付として解釈できません: {value!r} (YYYY-MM-DD 形式)")


def _datetime(value) -> Optional[_dt.datetime]:
    if value in (None, ""):
        return None
    if isinstance(value, _dt.datetime):
        return value
    if isinstance(value, _dt.date):
        return _dt.datetime(value.year, value.month, value.day)
    text = str(value).strip().replace("/", "-")
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M", "%Y-%m-%d"):
        try:
            return _dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    raise ProjectError(f"日時として解釈できません: {value!r}")


def _int(value, field: str) -> Optional[int]:
    if value in (None, ""):
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        raise ProjectError(f"{field} は整数である必要があります: {value!r}")


def _float(value, field: str) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ProjectError(f"{field} は数値である必要があります: {value!r}")


def _progress(value, index: int) -> Optional[float]:
    if value in (None, ""):
        return None
    text = str(value).strip()
    try:
        number = float(text.rstrip("%")) / 100.0 if text.endswith("%") else float(text)
    except ValueError:
        raise ProjectError(f"tasks[{index}].progress が不正です: {value!r}")
    if number > 1.0 and number <= 100.0:
        number /= 100.0
    if not 0.0 <= number <= 1.0:
        raise ProjectError(f"tasks[{index}].progress は 0〜1 (または 0%〜100%) です: {value!r}")
    return round(number, 4)


def _no(value) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()
