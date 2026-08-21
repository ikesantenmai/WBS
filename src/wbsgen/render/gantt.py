"""ガントチャート図形の組み立て。

タイムライン・行座標・タスクから :class:`~wbsgen.render.drawing.Drawing` を作る。
"""

from __future__ import annotations

import datetime as _dt
from typing import Dict, List, Optional

from .. import style
from ..model import KIND_MILESTONE, KIND_SUMMARY, MILESTONE_SHAPES, Project, Task
from ..scheduling import inazuma_points
from .drawing import Drawing, Geometry, Shape
from .timeline import Timeline

# 行内の縦位置 (0.0=上端, 1.0=下端)
PLAN_TOP_WITH_RESULT = 0.12
PLAN_BOTTOM_WITH_RESULT = 0.50
RESULT_TOP = 0.54
RESULT_BOTTOM = 0.92
PLAN_TOP_ALONE = 0.22
PLAN_BOTTOM_ALONE = 0.78


def _label(task: Task) -> str:
    """図形名に使う識別子 (項番と項目名)。"""
    return f"{task.no}:{task.name}" if task.no else task.name


class GanttBuilder:
    """タスク行からガントチャートの図形を作る。"""

    def __init__(self, project: Project, timeline: Timeline, geometry: Geometry,
                 first_chart_col: int, header_row: int, last_row: int, calendar):
        self.project = project
        self.timeline = timeline
        self.geo = geometry
        self.first_col = first_chart_col
        self.header_row = header_row      # 0 始まり: 日付ヘッダの行
        self.last_row = last_row          # 0 始まり: チャート領域の最終行
        self.calendar = calendar
        self.drawing = Drawing()
        self._colors = self._member_colors()
        self._bar_span: Dict[int, tuple] = {}   # task row -> (左座標, 右座標)

    # ------------------------------------------------------------------
    def _member_colors(self) -> Dict[str, str]:
        colors = {}
        palette = list(style.MEMBER_PALETTE)
        for i, member in enumerate(self.project.members):
            color = (member.color or palette[i % len(palette)]).lstrip("#").upper()
            colors[member.name] = color
        return colors

    def color_for(self, task: Task) -> str:
        if not self.project.chart.show.member_color:
            return style.MEMBER_PALETTE[0]
        color = self._colors.get(task.member)
        if color:
            return color
        # 未登録の担当者には名前から安定したパレット色を割り当てる
        if task.member:
            idx = sum(ord(c) for c in task.member) % len(style.MEMBER_PALETTE)
            return style.MEMBER_PALETTE[idx]
        return style.MEMBER_PALETTE[0]

    # ------------------------------------------------------------------
    def build(self) -> Drawing:
        show = self.project.chart.show
        for task in self.project.tasks:
            if task.row <= 0:
                continue
            if task.kind == KIND_MILESTONE:
                self._milestone(task)
            elif task.kind == KIND_SUMMARY:
                self._summary_bar(task)
            else:
                self._task_bars(task)
        if show.predecessor_line:
            self._predecessor_lines()
        if show.now_line:
            self._now_line()
        if show.inazuma_line:
            self._inazuma_line()
        return self.drawing.finalize(self.geo)

    # ------------------------------------------------------------------
    def _rect(self, left: float, right: float, row0: int, top: float, bottom: float,
              **kwargs) -> Optional[Shape]:
        """列座標 ``left``〜``right``・行 ``row0`` にバーを置く。"""
        if right <= left:
            right = left + 0.06   # 0 幅を避ける (1 日分に満たないバー)
        frm = self.geo.anchor(left, row0, top, self.first_col)
        to = self.geo.anchor(right, row0, bottom, self.first_col)
        return self.drawing.add(Shape(frm=frm, to=to, **kwargs))

    def _bar_range(self, start: _dt.date, end: _dt.date):
        """予定/実績の期間をチャート座標に変換する。範囲外なら ``None``。"""
        if start is None or end is None:
            return None
        if not self.timeline.in_range(start, end):
            return None
        left = self.timeline.position(self.timeline.clamp(start))
        right = self.timeline.position(self.timeline.clamp(end), end_of_day=True)
        return left, right

    # ------------------------------------------------------------------
    def _task_bars(self, task: Task) -> None:
        show = self.project.chart.show
        row0 = task.row - 1
        color = self.color_for(task)
        dark = style.shade(color, 0.55)
        light = style.shade(color, 1.55)

        with_result = show.results and task.actual_start is not None
        if with_result:
            p_top, p_bottom = PLAN_TOP_WITH_RESULT, PLAN_BOTTOM_WITH_RESULT
        else:
            p_top, p_bottom = PLAN_TOP_ALONE, PLAN_BOTTOM_ALONE

        span = self._bar_range(task.start, task.end)
        if span:
            left, right = span
            self._rect(
                left, right, row0, p_top, p_bottom,
                preset="rect", gradient=(dark, light),
                line_color=style.shade(color, 0.45), line_width_px=0.75,
                name=f"plan-{_label(task)}",
            )
            self._bar_span[task.row] = (left, right)
            if show.progress and task.progress:
                width = (right - left) * min(task.progress, 1.0)
                if width > 0:
                    self._rect(
                        left, left + width, row0, p_top + 0.10, p_bottom - 0.10,
                        preset="rect", fill=style.shade(color, 0.35),
                        name=f"progress-{_label(task)}",
                    )

        if with_result:
            actual_end = task.actual_end or self._elapsed_end(task)
            span = self._bar_range(task.actual_start, actual_end)
            if span:
                left, right = span
                self._rect(
                    left, right, row0, RESULT_TOP, RESULT_BOTTOM,
                    preset="rect", fill=style.shade(color, 0.95),
                    line_color=style.shade(color, 0.45), line_width_px=0.75,
                    alpha=100000 if task.is_complete else 70000,
                    name=f"actual-{_label(task)}",
                )

    def _elapsed_end(self, task: Task) -> _dt.date:
        """未完了タスクの実績バー右端 (現在日まで)。"""
        base = self.project.chart.base_date or _dt.date.today()
        if task.actual_days:
            return self.calendar.end_date(task.actual_start, task.actual_days)
        return max(task.actual_start, base)

    # ------------------------------------------------------------------
    def _summary_bar(self, task: Task) -> None:
        span = self._bar_range(task.start, task.end)
        if not span:
            return
        left, right = span
        color = self.color_for(task)
        label = task.name
        if task.start and task.end:
            label = f"{task.name} ({task.start.month}/{task.start.day}〜{task.end.month}/{task.end.day})"
        self._rect(
            left, right, task.row - 1, 0.15, 0.85,
            preset="chevron", adjust=0.06,
            gradient=(style.shade(color, 0.40), style.shade(color, 1.30)),
            line_color=None, text=label, text_color="FFFFFF", text_bold=True,
            text_align="ctr", name=f"summary-{_label(task)}",
        )
        self._bar_span[task.row] = (left, right)

    # ------------------------------------------------------------------
    def _milestone(self, task: Task) -> None:
        day = task.start or task.actual_start
        if day is None or not self.timeline.in_range(day, day):
            return
        pos = self.timeline.position(self.timeline.clamp(day))
        row0 = task.row - 1
        preset = MILESTONE_SHAPES.get(task.shape or "diamond", "diamond")
        color = self.color_for(task)
        # 記号は日付位置を中心に、行高と同じ幅で配置する
        half = self._marker_half_width(row0)
        self._rect(
            pos - half, pos + half, row0, 0.10, 0.90,
            preset=preset, fill=style.shade(color, 0.75),
            line_color=style.shade(color, 0.40), line_width_px=1.0,
            name=f"milestone-{_label(task)}",
        )
        self._bar_span[task.row] = (pos - half, pos + half)

    def _marker_half_width(self, row0: int) -> float:
        """行高と同じ大きさの記号になる列座標の半幅。"""
        col_px = self.geo.col_width_emu(self.first_col) / 9525.0
        row_px = self.geo.row_height_emu(row0) / 12700.0 * (96 / 72.0)
        return max(row_px / col_px / 2.0, 0.12)

    # ------------------------------------------------------------------
    def _predecessor_lines(self) -> None:
        by_no = {str(t.no): t for t in self.project.tasks if t.no not in ("", None)}
        for task in self.project.tasks:
            if not task.predecessor or task.row <= 0:
                continue
            for key in str(task.predecessor).replace("、", ",").split(","):
                pred = by_no.get(key.strip())
                if pred is None or pred.row <= 0:
                    continue
                self._link_line(pred, task)

    def _link_line(self, pred: Task, succ: Task) -> None:
        src = self._bar_span.get(pred.row)
        dst = self._bar_span.get(succ.row)
        if not src or not dst:
            return
        x1, y1 = src[1], 0.5
        x2, y2 = dst[0], 0.5
        row1, row2 = pred.row - 1, succ.row - 1
        # 縦→横の 2 本で L 字に結ぶ (右向き矢印を後段に付ける)
        mid = x1 + 0.25
        self._segment(x1, row1, y1, mid, row1, y1, arrow=False)
        self._segment(mid, row1, y1, mid, row2, y2, arrow=False)
        self._segment(mid, row2, y2, max(x2, mid), row2, y2, arrow=True)

    def _segment(self, x1: float, r1: int, f1: float, x2: float, r2: int, f2: float,
                 arrow: bool = False, color: str = style.C_LINK,
                 width: float = 0.75, dash: Optional[str] = None) -> None:
        """任意の 2 点を結ぶ直線図形を追加する。

        ``line`` プリセットは外接矩形の左上から右下へ引かれるため、
        右上→左下の線は ``flipH`` で反転させる。
        """
        upper, lower = ((x1, r1, f1), (x2, r2, f2))
        if (r1 + f1) > (r2 + f2):
            upper, lower = lower, upper
        frm = self.geo.anchor(min(x1, x2), upper[1], upper[2], self.first_col)
        to = self.geo.anchor(max(x1, x2), lower[1], lower[2], self.first_col)
        self.drawing.add(Shape(
            frm=frm, to=to, preset="line", fill=None,
            line_color=color, line_width_px=width, dash=dash,
            arrow_head=arrow, flip_h=upper[0] > lower[0], name="link",
        ))

    # ------------------------------------------------------------------
    def _now_line(self) -> None:
        base = self.project.chart.base_date or _dt.date.today()
        if not self.timeline.in_range(base, base):
            return
        pos = self.timeline.position(base)
        frm = self.geo.anchor(pos, self.header_row, 0.0, self.first_col)
        to = self.geo.anchor(pos, self.last_row, 1.0, self.first_col)
        self.drawing.add(Shape(
            frm=frm, to=to, preset="line", fill=None,
            line_color=style.C_NOWLINE, line_width_px=1.25, dash="dash",
            name="now-line",
        ))

    # ------------------------------------------------------------------
    def _inazuma_line(self) -> None:
        base = self.project.chart.base_date or _dt.date.today()
        if not self.timeline.in_range(base, base):
            return
        now_pos = self.timeline.position(base)
        points = []
        for task, achieved in inazuma_points(self.project, self.calendar, base):
            if task.row <= 0:
                continue
            if not self.timeline.in_range(achieved, achieved):
                continue
            pos = self.timeline.position(self.timeline.clamp(achieved),
                                         end_of_day=task.is_complete)
            points.append((pos, task.row - 1))
        if not points:
            return
        # 先頭・末尾は現在日線に接続する
        prev = (now_pos, self.header_row)
        for pos, row in points:
            self._segment(prev[0], prev[1], 1.0 if prev[1] == self.header_row else 0.5,
                          pos, row, 0.5, color=style.C_INAZUMA, width=1.25)
            prev = (pos, row)
        self._segment(prev[0], prev[1], 0.5, now_pos, prev[1], 0.5,
                      color=style.C_INAZUMA, width=1.25)
