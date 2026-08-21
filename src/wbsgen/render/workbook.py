"""Excel ブックの生成。

添付の WBS と同じ 4 シート構成 (スケジュール / 担当者一覧 / 標準工程 / 設定) を作る。
"""

from __future__ import annotations

import datetime as _dt
from pathlib import Path
from typing import Dict, List

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .. import style
from ..model import KIND_SUMMARY, Project, Task
from ..scheduling import group_effort, resolve_project
from .drawing import Geometry
from .gantt import GanttBuilder
from .inject import inject_drawing
from .timeline import Timeline

SHEET_PLAN = "スケジュール"
SHEET_MEMBER = "担当者一覧"
SHEET_PROCESS = "標準工程"
SHEET_CONFIG = "設定"

ROW_TITLE = 1
ROW_BAND = 3        # 年月の帯
ROW_HEADER = 4      # 見出し / 日付
ROW_FIRST_TASK = 5

#: 表側の見出し (列文字, 見出し, 幅キー)
TABLE_HEADERS = [
    (style.COL_GROUP, "大項目"),
    (style.COL_SUBGROUP, "中項目"),
    (style.COL_NO, "項番"),
    (style.COL_NAME, "項目"),
    (style.COL_START, "開始日"),
    (style.COL_DAYS, "日数"),
    (style.COL_END, "終了日"),
    (style.COL_ASTART, "開始"),
    (style.COL_ADAYS, "日数"),
    (style.COL_AEND, "終了"),
    (style.COL_DELAY, "遅れ"),
    (style.COL_PROGRESS, "進捗"),
    (style.COL_EFFORT, "工数"),
    (style.COL_PRED, "先行"),
    (style.COL_SHAPE, ""),
    (style.COL_MEMBER, "担当"),
    (style.COL_STATUS, "状態"),
]


class WorkbookRenderer:
    """:class:`~wbsgen.model.Project` から xlsx を書き出す。"""

    def __init__(self, project: Project):
        self.project = project
        self.calendar = resolve_project(project)
        chart = project.chart
        self.timeline = Timeline(chart.start, chart.period_days, chart.unit, self.calendar)
        self.wb = Workbook()
        self._row_heights: Dict[int, float] = {}
        self._chart_col_px = style.CHART_COL_WIDTH_PX[chart.unit]

    # ------------------------------------------------------------------
    def save(self, path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        ws = self.wb.active
        ws.title = SHEET_PLAN
        self._build_plan(ws)
        self._build_members(self.wb.create_sheet(SHEET_MEMBER))
        self._build_process(self.wb.create_sheet(SHEET_PROCESS))
        self._build_config(self.wb.create_sheet(SHEET_CONFIG))
        self.wb.save(path)

        drawing = self._build_gantt()
        if len(drawing):
            inject_drawing(path, "xl/worksheets/sheet1.xml", drawing.to_xml())
        return path

    # ==================================================================
    # スケジュールシート
    # ==================================================================
    def _build_plan(self, ws: Worksheet) -> None:
        self._layout_columns(ws)
        self._title(ws)
        self._table_header(ws)
        self._timeline_header(ws)
        self._task_rows(ws)
        self._freeze(ws)

    def _layout_columns(self, ws: Worksheet) -> None:
        for letter, px in style.COLUMN_WIDTHS_PX.items():
            ws.column_dimensions[letter].width = style.px_to_width(px)
        hidden = self._hidden_columns()
        for letter in hidden:
            ws.column_dimensions[letter].hidden = True
        chart_width = style.px_to_width(self._chart_col_px)
        for i in range(len(self.timeline)):
            letter = get_column_letter(style.COL_CHART_FIRST + i)
            ws.column_dimensions[letter].width = chart_width

    def _hidden_columns(self) -> List[str]:
        show = self.project.chart.show
        hidden = [style.COL_SHAPE]
        if not show.start_date:
            hidden.append(style.COL_START)
        if not show.end_date:
            hidden.append(style.COL_END)
        if not show.results:
            hidden += [style.COL_ASTART, style.COL_ADAYS, style.COL_AEND, style.COL_DELAY]
        if not show.progress:
            pass  # 進捗列は実績と一緒に常に出す (元ツールと同じ)
        if not show.manpower:
            hidden.append(style.COL_EFFORT)
        if not show.predecessor:
            hidden.append(style.COL_PRED)
        if not show.status:
            hidden.append(style.COL_STATUS)
        return hidden

    # ------------------------------------------------------------------
    def _title(self, ws: Worksheet) -> None:
        self._set_row_height(ws, ROW_TITLE, style.ROW_HEIGHT_TITLE)
        cell = ws.cell(row=ROW_TITLE, column=2, value=self.project.title)
        cell.font = style.font(12, bold=True, color=style.C_TITLE_FONT)
        cell.alignment = style.ALIGN_LEFT

        updated = self.project.updated_at or _dt.datetime.now()
        stamp = ws.cell(row=ROW_TITLE, column=style.COL_CHART_FIRST - 3, value=updated)
        stamp.number_format = style.FMT_UPDATED
        stamp.font = style.font(9, color="808080")
        stamp.alignment = style.ALIGN_RIGHT

        self._set_row_height(ws, 2, 6.0)

    def _table_header(self, ws: Worksheet) -> None:
        self._set_row_height(ws, ROW_BAND, style.ROW_HEIGHT_HEADER)
        self._set_row_height(ws, ROW_HEADER, style.ROW_HEIGHT_HEADER)

        # 予定 / 実績のグループ見出し
        ws.merge_cells(f"{style.COL_START}{ROW_BAND}:{style.COL_END}{ROW_BAND}")
        ws.merge_cells(f"{style.COL_ASTART}{ROW_BAND}:{style.COL_PROGRESS}{ROW_BAND}")
        for letter, text in ((style.COL_START, "予定"), (style.COL_ASTART, "実績")):
            cell = ws[f"{letter}{ROW_BAND}"]
            cell.value = text
            self._style_header(cell)

        for letter in ("B", "C", "D", "E", "N", "O", "P", "Q", "R"):
            self._style_header(ws[f"{letter}{ROW_BAND}"])

        for letter, text in TABLE_HEADERS:
            cell = ws[f"{letter}{ROW_HEADER}"]
            cell.value = text
            self._style_header(cell)

    def _style_header(self, cell) -> None:
        cell.fill = style.fill(style.C_HEADER)
        cell.font = style.font(bold=False)
        cell.alignment = style.ALIGN_CENTER
        cell.border = style.BORDER_CELL

    # ------------------------------------------------------------------
    def _timeline_header(self, ws: Worksheet) -> None:
        first = style.COL_CHART_FIRST
        # 上段: 年月の帯
        for start, span, text in self.timeline.band_labels():
            if span <= 0:
                continue
            left = get_column_letter(first + start)
            right = get_column_letter(first + start + span - 1)
            if span > 1:
                ws.merge_cells(f"{left}{ROW_BAND}:{right}{ROW_BAND}")
            cell = ws[f"{left}{ROW_BAND}"]
            cell.value = text
            cell.fill = style.fill(style.C_TIMELINE)
            cell.font = style.font(bold=True)
            cell.alignment = style.ALIGN_CENTER
            cell.border = style.BORDER_CELL

        # 下段: 日付/週/月ラベル
        for col in self.timeline.columns:
            cell = ws.cell(row=ROW_HEADER, column=first + col.index)
            cell.value = self.timeline.column_label(col)
            cell.number_format = style.FMT_TEXT
            rest = self.timeline.is_rest_column(col)
            cell.fill = style.fill(style.C_HOLIDAY if rest else style.C_TIMELINE)
            cell.font = style.font(8, color="C00000" if rest else "000000")
            cell.alignment = style.ALIGN_CENTER
            cell.border = style.BORDER_CELL

    # ------------------------------------------------------------------
    def _task_rows(self, ws: Worksheet) -> None:
        show = self.project.chart.show
        row = ROW_FIRST_TASK
        prev_group = None
        prev_subgroup = None

        for task in self.project.tasks:
            if prev_group is not None and task.group and task.group != prev_group:
                self._group_footer(ws, row, prev_group)
                row += 1
            task.row = row
            self._write_task(ws, row, task,
                             new_group=task.group != prev_group,
                             new_subgroup=(task.group, task.subgroup) != (prev_group, prev_subgroup))
            prev_group, prev_subgroup = task.group, task.subgroup
            row += 1

        if prev_group is not None:
            self._group_footer(ws, row, prev_group)
            row += 1
        self._last_row = row - 1
        self._chart_grid(ws, ROW_FIRST_TASK, self._last_row)

    def _write_task(self, ws: Worksheet, row: int, task: Task,
                    new_group: bool, new_subgroup: bool) -> None:
        self._set_row_height(ws, row, style.ROW_HEIGHT_TASK)
        is_group_row = task.kind == KIND_SUMMARY or new_group
        base_fill = style.C_GROUP if is_group_row else style.C_WHITE

        def put(letter, value, number_format=None, align=None, fill_color=None,
                font=None):
            cell = ws[f"{letter}{row}"]
            cell.value = value
            if number_format:
                cell.number_format = number_format
            cell.alignment = align or style.ALIGN_CENTER
            cell.fill = style.fill(fill_color or base_fill)
            cell.font = font or style.font(bold=is_group_row and letter == style.COL_GROUP)
            cell.border = style.BORDER_CELL
            return cell

        put(style.COL_GROUP, task.group if new_group else None, align=style.ALIGN_LEFT)
        put(style.COL_SUBGROUP, task.subgroup if new_subgroup else None, align=style.ALIGN_LEFT)
        put(style.COL_NO, task.no, number_format=style.FMT_TEXT)
        put(style.COL_NAME, task.name, align=style.ALIGN_NAME)

        plan_font = style.font(color=style.C_PLAN_FONT)
        plan_fill = style.C_PLAN_CELL if not is_group_row else base_fill
        put(style.COL_START, task.start, style.FMT_DATE, fill_color=plan_fill, font=plan_font)
        put(style.COL_DAYS, task.days, style.FMT_DAYS, fill_color=plan_fill, font=plan_font)
        put(style.COL_END, task.end, style.FMT_DATE, fill_color=plan_fill, font=plan_font)

        put(style.COL_ASTART, task.actual_start, style.FMT_DATE)
        put(style.COL_ADAYS, task.actual_days, style.FMT_DAYS)
        put(style.COL_AEND, task.actual_end, style.FMT_DATE)

        delay_font = style.font(color="FF0000", bold=True) if task.delay else style.font()
        put(style.COL_DELAY, task.delay, style.FMT_DAYS, font=delay_font)
        put(style.COL_PROGRESS, task.progress, style.FMT_PERCENT, font=delay_font)
        put(style.COL_EFFORT, task.effort, style.FMT_EFFORT)
        put(style.COL_PRED, task.predecessor, number_format=style.FMT_TEXT)
        put(style.COL_SHAPE, None)
        put(style.COL_MEMBER, task.member, align=style.ALIGN_LEFT)

        bg, fg, bold = style.status_style(task.status or "-")
        put(style.COL_STATUS, task.status or "-", fill_color=bg,
            font=style.font(color=fg, bold=bold))

        if task.comment:
            ws[f"{style.COL_NAME}{row}"].comment = _comment(task.comment)

    def _group_footer(self, ws: Worksheet, row: int, group: str) -> None:
        """大項目の区切り行 (工数合計)。"""
        self._set_row_height(ws, row, style.ROW_HEIGHT_SPACER)
        effort = group_effort(self.project, group)
        cell = ws.cell(row=row, column=2, value=f"[ {effort:g} 人日 ]" if effort else None)
        cell.font = style.font(7, color="7F7F7F")
        cell.alignment = style.ALIGN_RIGHT
        for col in range(2, style.COL_CHART_FIRST):
            ws.cell(row=row, column=col).fill = style.fill(style.C_GROUP)

    # ------------------------------------------------------------------
    def _chart_grid(self, ws: Worksheet, first_row: int, last_row: int) -> None:
        """チャート領域の罫線・休日列の色分け。"""
        first = style.COL_CHART_FIRST
        for col in self.timeline.columns:
            rest = self.timeline.is_rest_column(col)
            sat = self.timeline.is_saturday_column(col)
            color = style.C_HOLIDAY if rest and not sat else (
                style.C_SATURDAY if sat else style.C_WHITE)
            for row in range(first_row, last_row + 1):
                cell = ws.cell(row=row, column=first + col.index)
                cell.fill = style.fill(color)
                cell.border = style.BORDER_CHART

    def _freeze(self, ws: Worksheet) -> None:
        ws.freeze_panes = f"{get_column_letter(style.COL_CHART_FIRST)}{ROW_FIRST_TASK}"
        ws.sheet_view.showGridLines = False

    def _set_row_height(self, ws: Worksheet, row: int, height: float) -> None:
        ws.row_dimensions[row].height = height
        self._row_heights[row] = height

    # ==================================================================
    # ガントチャート図形
    # ==================================================================
    def _build_gantt(self):
        chart_px = self._chart_col_px
        table_px = style.COLUMN_WIDTHS_PX

        def col_px(col0: int) -> float:
            letter = get_column_letter(col0 + 1)
            if letter in table_px:
                return table_px[letter]
            return chart_px

        def row_pt(row0: int) -> float:
            return self._row_heights.get(row0 + 1, style.ROW_HEIGHT_TASK)

        geometry = Geometry(col_px, row_pt)
        builder = GanttBuilder(
            self.project, self.timeline, geometry,
            first_chart_col=style.COL_CHART_FIRST - 1,
            header_row=ROW_HEADER - 1,
            calendar=self.calendar,
        )
        return builder.build()

    # ==================================================================
    # 担当者一覧シート
    # ==================================================================
    def _build_members(self, ws: Worksheet) -> None:
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = style.px_to_width(25)
        for letter, px in (("B", 130), ("C", 60), ("D", 90), ("E", 90), ("F", 90), ("G", 110)):
            ws.column_dimensions[letter].width = style.px_to_width(px)

        ws["B2"] = "◆担当者一覧"
        ws["B2"].font = style.font(11, bold=True, color=style.C_TITLE_FONT)

        headers = ["担当", "色", "稼働曜日", "個別休日", "個別出勤", "サンプル"]
        for i, text in enumerate(headers):
            cell = ws.cell(row=3, column=2 + i, value=text)
            self._style_header(cell)

        palette = list(style.MEMBER_PALETTE)
        for i, member in enumerate(self.project.members):
            row = 4 + i
            color = (member.color or palette[i % len(palette)]).lstrip("#").upper()
            values = [
                member.name,
                f"#{color}",
                "/".join(member.workdays) if member.workdays else "(既定)",
                _join_dates(member.holidays),
                _join_dates(member.extra_workdays),
                None,
            ]
            for j, value in enumerate(values):
                cell = ws.cell(row=row, column=2 + j, value=value)
                cell.border = style.BORDER_CELL
                cell.alignment = style.ALIGN_LEFT
                cell.font = style.font()
            sample = ws.cell(row=row, column=7)
            sample.fill = style.fill(style.shade(color, 1.4))
            ws.cell(row=row, column=3).fill = style.fill(color)
            ws.cell(row=row, column=3).font = style.font(color="FFFFFF")

    # ==================================================================
    # 標準工程シート
    # ==================================================================
    def _build_process(self, ws: Worksheet) -> None:
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = style.px_to_width(25)
        for letter, px in (("B", 110), ("C", 110), ("D", 52), ("E", 240),
                           ("F", 56), ("G", 48), ("H", 52), ("I", 72)):
            ws.column_dimensions[letter].width = style.px_to_width(px)

        ws["B2"] = "◆標準工程"
        ws["B2"].font = style.font(11, bold=True, color=style.C_TITLE_FONT)

        headers = ["大項目", "中項目", "項番", "項目", "日数", "工数", "先行", "担当"]
        for i, text in enumerate(headers):
            self._style_header(ws.cell(row=4, column=2 + i, value=text))

        source = self.project.standard_process or self.project.tasks
        for i, task in enumerate(source):
            row = 5 + i
            values = [task.group, task.subgroup, task.no, task.name,
                      task.days, task.effort, task.predecessor, task.member]
            for j, value in enumerate(values):
                cell = ws.cell(row=row, column=2 + j, value=value)
                cell.border = style.BORDER_CELL
                cell.font = style.font()
                cell.alignment = style.ALIGN_LEFT if j in (0, 1, 3, 7) else style.ALIGN_CENTER
            ws.cell(row=row, column=6).number_format = style.FMT_DAYS
            ws.cell(row=row, column=7).number_format = style.FMT_EFFORT

    # ==================================================================
    # 設定シート
    # ==================================================================
    def _build_config(self, ws: Worksheet) -> None:
        ws.sheet_view.showGridLines = False
        ws.column_dimensions["A"].width = style.px_to_width(25)
        for letter, px in (("B", 170), ("C", 150), ("E", 170), ("F", 120)):
            ws.column_dimensions[letter].width = style.px_to_width(px)

        chart = self.project.chart
        show = chart.show
        blocks = [
            ("◆チャート表示設定", [
                ("チャート表示開始日", self.timeline.start),
                ("チャート表示期間(日)", chart.period_days),
                ("チャート表示単位", {"day": "日単位", "week": "週単位",
                                      "month": "月単位"}[chart.unit]),
                ("表示基準日(現在日)", chart.base_date),
                ("1人月=", chart.man_month_days),
            ]),
            ("◆タスク表示設定", [
                ("実績表示", _onoff(show.results)),
                ("進捗率表示", _onoff(show.progress)),
                ("工数表示", _onoff(show.manpower)),
                ("状態表示", _onoff(show.status)),
                ("開始日表示", _onoff(show.start_date)),
                ("終了日表示", _onoff(show.end_date)),
                ("先行タスク線表示", _onoff(show.predecessor_line)),
                ("現在日線表示", _onoff(show.now_line)),
                ("イナズマ線表示", _onoff(show.inazuma_line)),
                ("担当色", _onoff(show.member_color)),
            ]),
            ("◆状態設定", [
                ("完了", "完了"),
                ("実行中", "実行中"),
                ("実行中残り", f"残り %d 日 ({chart.thresholds.exec_remain_days}日以内)"),
                ("遅れ", "遅れ %d 日"),
                ("予定間近", f"あと %d 日 ({chart.thresholds.start_near_days}日以内)"),
            ]),
            ("◆作業日設定", [
                (name, "出" if index in self.calendar.workdays else "休")
                for index, name in enumerate(
                    ("月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"))
            ]),
        ]

        row = 2
        for title, entries in blocks:
            cell = ws.cell(row=row, column=2, value=title)
            cell.font = style.font(11, bold=True, color=style.C_TITLE_FONT)
            row += 1
            for label, value in entries:
                ws.cell(row=row, column=2, value=label).font = style.font()
                target = ws.cell(row=row, column=3, value=value)
                target.font = style.font()
                target.alignment = style.ALIGN_LEFT
                if isinstance(value, _dt.date):
                    target.number_format = "yyyy/mm/dd"
                ws.cell(row=row, column=2).border = style.BORDER_CELL
                target.border = style.BORDER_CELL
                row += 1
            row += 1

        # 休日一覧
        ws.cell(row=2, column=5, value="◆休日設定").font = style.font(
            11, bold=True, color=style.C_TITLE_FONT)
        holidays = sorted(d for d in self.calendar.holidays
                          if self.timeline.start <= d <= self.timeline.end)
        for i, day in enumerate(holidays):
            cell = ws.cell(row=3 + i, column=5, value=day)
            cell.number_format = "yyyy/mm/dd"
            cell.font = style.font()
            cell.border = style.BORDER_CELL


# ----------------------------------------------------------------------
def _onoff(flag: bool) -> str:
    return "ON" if flag else "OFF"


def _join_dates(days) -> str:
    return ", ".join(d.strftime("%m/%d") for d in days) if days else ""


def _comment(text: str):
    from openpyxl.comments import Comment

    comment = Comment(text, "wbsgen")
    comment.width = 260
    comment.height = 90
    return comment
