"""WBS のデータモデル。

YAML/JSON/CSV から読み込んだ内容を保持するだけのプレーンなデータクラス群。
日数・状態などの導出は :mod:`wbsgen.scheduling` が行う。
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Optional

#: タスク種別
KIND_TASK = "task"          # 通常タスク (バー)
KIND_SUMMARY = "summary"    # 工程 (サマリバー)
KIND_MILESTONE = "milestone"  # マイルストーン (記号)

VALID_KINDS = (KIND_TASK, KIND_SUMMARY, KIND_MILESTONE)

#: 表示単位
UNIT_DAY = "day"
UNIT_WEEK = "week"
UNIT_MONTH = "month"
VALID_UNITS = (UNIT_DAY, UNIT_WEEK, UNIT_MONTH)

#: マイルストーン記号 (元ツールの図形種類に対応)
MILESTONE_SHAPES = {
    "diamond": "diamond",
    "triangle": "triangle",
    "circle": "ellipse",
    "star": "star5",
    "arrow": "rightArrow",
    "chevron": "chevron",
}


@dataclass
class Member:
    """担当者 (チーム)。"""

    name: str
    color: str = "#4472C4"
    #: 休日区分。既定カレンダーと別の稼働日を持つ担当者に使う。
    workdays: Optional[list] = None
    holidays: list = field(default_factory=list)
    extra_workdays: list = field(default_factory=list)


@dataclass
class Task:
    """WBS の 1 行。"""

    name: str
    group: str = ""            # 大項目
    subgroup: str = ""         # 中項目
    no: str = ""               # 項番
    kind: str = KIND_TASK
    start: Optional[_dt.date] = None       # 予定開始日
    days: Optional[int] = None             # 予定日数 (稼働日)
    end: Optional[_dt.date] = None         # 予定終了日 (未指定なら導出)
    actual_start: Optional[_dt.date] = None
    actual_days: Optional[int] = None
    actual_end: Optional[_dt.date] = None
    progress: Optional[float] = None       # 0.0〜1.0
    effort: Optional[float] = None         # 工数 (人日)
    predecessor: str = ""                  # 先行タスクの項番
    member: str = ""
    comment: str = ""
    shape: str = ""                        # マイルストーン記号 (kind=milestone)

    # --- 導出値 (scheduling が設定する) ---
    delay: Optional[int] = field(default=None, compare=False)
    status: str = field(default="", compare=False)
    row: int = field(default=0, compare=False)

    @property
    def is_bar(self) -> bool:
        return self.kind in (KIND_TASK, KIND_SUMMARY)

    @property
    def is_complete(self) -> bool:
        return self.progress is not None and self.progress >= 1.0

    @property
    def has_plan(self) -> bool:
        return self.start is not None and self.end is not None


@dataclass
class DisplayFlags:
    """表示 ON/OFF 設定 (元ツールの「設定」シート相当)。"""

    results: bool = True            # 実績表示
    progress: bool = False          # 進捗率表示
    manpower: bool = True           # 工数表示
    status: bool = True             # 状態表示
    start_date: bool = True         # 予定開始日列
    end_date: bool = True           # 予定終了日列
    predecessor: bool = True        # 先行列
    predecessor_line: bool = True   # 先行タスク線
    now_line: bool = True           # 現在日線
    inazuma_line: bool = True       # イナズマ線
    member_color: bool = True       # 担当色
    comment: bool = True            # コメント


@dataclass
class StatusThresholds:
    """状態表示のしきい値。"""

    exec_remain_days: int = 5   # 「残り %d 日」に切り替わる残稼働日数
    start_near_days: int = 10   # 「あと %d 日」に切り替わる着手前稼働日数


@dataclass
class ChartConfig:
    """チャート表示設定。"""

    start: Optional[_dt.date] = None      # チャート表示開始日 (未指定なら自動)
    period_days: int = 360                # チャート表示期間 (暦日)
    unit: str = UNIT_WEEK                 # 表示単位
    base_date: Optional[_dt.date] = None  # 現在日 (未指定なら実行日)
    man_month_days: int = 20              # 1 人月の日数
    thresholds: StatusThresholds = field(default_factory=StatusThresholds)
    show: DisplayFlags = field(default_factory=DisplayFlags)

    def __post_init__(self):
        if self.unit not in VALID_UNITS:
            raise ValueError(f"表示単位が不正です: {self.unit!r} (day/week/month)")


@dataclass
class CalendarConfig:
    """稼働日設定。"""

    workdays: list = field(default_factory=lambda: ["mon", "tue", "wed", "thu", "fri"])
    japanese_holidays: bool = True
    holidays: list = field(default_factory=list)
    extra_workdays: list = field(default_factory=list)


@dataclass
class Project:
    """WBS プロジェクト全体。"""

    title: str = "プロジェクト スケジュール"
    updated_at: Optional[_dt.datetime] = None
    chart: ChartConfig = field(default_factory=ChartConfig)
    calendar: CalendarConfig = field(default_factory=CalendarConfig)
    members: list = field(default_factory=list)
    tasks: list = field(default_factory=list)
    #: タスクの後ろに追加する空行の数 (記入用の枠)
    blank_rows: int = 0
    #: 標準工程シートに出力するテンプレートタスク
    standard_process: list = field(default_factory=list)

    def member(self, name: str) -> Optional[Member]:
        for m in self.members:
            if m.name == name:
                return m
        return None
