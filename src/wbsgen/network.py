"""ネットワーク図 (先行関係の依存グラフ) の組み立て。

「先行」欄に書かれた項番から、タスクの前後関係を矢印でつなぐ。
段 (:data:`level`) は先行タスクが無ければ 0、あれば先行の中で最も遅い段の
次にする (PERT 図と同じ、最長経路による並べ方)。

項番が無い行 (大項目だけのまとめ行など) は対象にしない。項番は
:mod:`wbsgen.daily` の整合性チェックと同じ規則で読む。
"""

from __future__ import annotations

from typing import Any, Dict, List

from .daily import predecessors
from .importer import Row


def build(rows: List[Row]) -> Dict[str, Any]:
    """行の一覧から、ネットワーク図のノードと矢印を組み立てる。"""
    tasks = [row for row in rows if str(row.no).strip()]
    by_no: Dict[str, Row] = {}
    for row in tasks:
        by_no.setdefault(str(row.no).strip(), row)

    level = _levels(tasks, by_no)
    lane = _lanes(tasks, level, by_no)
    nodes = [_node(row, level[row.row], lane[row.row]) for row in tasks]
    edges = _edges(tasks, by_no)
    return {"nodes": nodes, "edges": edges}


def _levels(tasks: List[Row], by_no: Dict[str, Row]) -> Dict[int, int]:
    """各行の段。循環参照は先行が無かったものとして打ち切る (無限再帰を避ける)。"""
    level: Dict[int, int] = {}

    def resolve(row: Row, visiting: set) -> int:
        if row.row in level:
            return level[row.row]
        if row.row in visiting:
            return 0
        visiting.add(row.row)
        lvl = 0
        for no in predecessors(row):
            before = by_no.get(no)
            if before is not None and before.row != row.row:
                lvl = max(lvl, resolve(before, visiting) + 1)
        visiting.discard(row.row)
        level[row.row] = lvl
        return lvl

    for row in tasks:
        resolve(row, set())
    return level


def _lanes(tasks: List[Row], level: Dict[int, int],
           by_no: Dict[str, Row]) -> Dict[int, int]:
    """段の中での上下位置。先行の位置の平均に近い側へ寄せて、矢印の交差を減らす
    (バリセンタ法)。先行の無い段はもともとの行の並び順のままにする。
    """
    order = {row.row: i for i, row in enumerate(tasks)}
    by_level: Dict[int, List[Row]] = {}
    for row in tasks:
        by_level.setdefault(level[row.row], []).append(row)

    sources_of: Dict[int, List[int]] = {}
    for row in tasks:
        for no in predecessors(row):
            before = by_no.get(no)
            if before is not None and before.row != row.row:
                sources_of.setdefault(row.row, []).append(before.row)

    lane: Dict[int, int] = {}
    for lvl in sorted(by_level):
        group = by_level[lvl]
        if lvl == 0:
            ordered = group
        else:
            def barycenter(row: Row) -> tuple:
                # まだレーンが決まっていない先行 (循環参照) は無視する
                sources = [lane[s] for s in sources_of.get(row.row, []) if s in lane]
                average = sum(sources) / len(sources) if sources else float(order[row.row])
                return (average, order[row.row])

            ordered = sorted(group, key=barycenter)
        for i, row in enumerate(ordered):
            lane[row.row] = i
    return lane


def _node(row: Row, level: int, lane: int) -> Dict[str, Any]:
    return {
        "row": row.row,
        "no": row.no,
        "name": row.name,
        "group": row.group,
        "subgroup": row.subgroup,
        "member": row.member,
        "start": _iso(row.start),
        "end": _iso(row.end),
        "progress": row.progress,
        "status": row.status,
        "delay": row.delay,
        "level": level,
        "lane": lane,
    }


def _edges(tasks: List[Row], by_no: Dict[str, Row]) -> List[Dict[str, Any]]:
    """先行のうち、この表に見つかったものだけを矢印にする (見つからない先行は
    :mod:`wbsgen.daily` の整合性チェックが別に知らせる)。"""
    edges = []
    for row in tasks:
        for no in predecessors(row):
            before = by_no.get(no)
            if before is None or before.row == row.row:
                continue
            edges.append({
                "from": before.row,
                "to": row.row,
                # 先行タスクの予定終了日より前に始まる予定になっているか
                "late": bool(before.end and row.start and before.end > row.start),
            })
    return edges


def _iso(value):
    return value.isoformat() if value else None


__all__ = ["build"]
