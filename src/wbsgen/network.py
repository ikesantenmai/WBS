"""ネットワーク図 (先行関係の依存グラフ) の組み立て。

「先行」欄に書かれた項番から、タスクの前後関係を図にする。2 とおりの
書き方に対応する。

- :func:`build` — **プレジデンス図** (PDM)。タスクそのものを箱にして、
  矢印で前後関係をつなぐ。
- :func:`build_arrow` — **アロー図** (ADM)。タスクを矢印にして、
  前後関係の合流点を丸 (イベント) で表す昔ながらの書き方。複数の先行を
  持つタスクは、ダミー矢印 (作業を表さない、順序をそろえるためだけの
  矢印) で合流させる。

どちらも、段 (:data:`level`) は先行が無ければ 0、あれば先行の中で最も
遅い段の次にする (最長経路による並べ方)。段の中の上下位置 (:data:`lane`)
は、先行の位置の平均に近い側へ寄せて矢印の交差を減らす (バリセンタ法)。

項番が無い行 (大項目だけのまとめ行など) は対象にしない。項番は
:mod:`wbsgen.daily` の整合性チェックと同じ規則で読む。
"""

from __future__ import annotations

from typing import Any, Dict, List, Tuple

from .daily import predecessors
from .importer import Row

Edge = Tuple[int, int]


def build(rows: List[Row]) -> Dict[str, Any]:
    """行の一覧から、プレジデンス図のノードと矢印を組み立てる。"""
    tasks = [row for row in rows if str(row.no).strip()]
    by_no = _by_no(tasks)

    level = _levels(tasks, by_no)
    lane = _lanes(tasks, level, by_no)
    times = _times(tasks, by_no, level)
    nodes = [_node(row, level[row.row], lane[row.row], times[row.row]) for row in tasks]
    edges = _edges(tasks, by_no, times)
    return {"nodes": nodes, "edges": edges}


def build_arrow(rows: List[Row]) -> Dict[str, Any]:
    """行の一覧から、アロー図のイベントと矢印を組み立てる。"""
    tasks = [row for row in rows if str(row.no).strip()]
    if not tasks:
        return {"events": [], "activities": []}
    by_no = _by_no(tasks)

    # イベント番号は、先行 (段が浅いほう) から順に振る
    task_level = _levels(tasks, by_no)
    order = {row.row: i for i, row in enumerate(tasks)}
    topo = sorted(tasks, key=lambda row: (task_level[row.row], order[row.row]))

    next_id = [0]

    def new_event() -> int:
        event = next_id[0]
        next_id[0] += 1
        return event

    start_event = new_event()
    head: Dict[int, int] = {}
    tail: Dict[int, int] = {}
    #: 同じ先行の組は同じ合流イベントを使う (ダミー矢印を重複させない)
    merge_of: Dict[frozenset, int] = {}
    dummies: List[Edge] = []

    for row in topo:
        preds = _valid_predecessors(row, by_no)
        if not preds:
            tail[row.row] = start_event
        elif len(preds) == 1:
            tail[row.row] = head.get(preds[0].row, start_event)
        else:
            key = frozenset(p.row for p in preds)
            merge_event = merge_of.get(key)
            if merge_event is None:
                merge_event = new_event()
                merge_of[key] = merge_event
                for p in preds:
                    dummies.append((head.get(p.row, start_event), merge_event))
            tail[row.row] = merge_event
        head[row.row] = new_event()

    # 後続の無いタスク (複数あれば、最後にダミーで 1 つのイベントへ束ねる)
    has_successor = {p.row for row in tasks for p in _valid_predecessors(row, by_no)}
    sinks = [row for row in tasks if row.row not in has_successor]
    if len(sinks) > 1:
        end_event = new_event()
        dummies += [(head[row.row], end_event) for row in sinks]
    else:
        end_event = head[sinks[0].row] if sinks else start_event

    activities = [_activity(row, tail[row.row], head[row.row], by_no) for row in tasks]
    activities += [{"from": a, "to": b, "dummy": True} for a, b in dummies]

    events = list(range(next_id[0]))
    edges = [(activity["from"], activity["to"]) for activity in activities]
    event_level = _graph_levels(events, edges)
    event_lane = _graph_lanes(events, event_level, edges)

    return {
        "events": [
            {"id": event, "level": event_level[event], "lane": event_lane[event]}
            for event in events
        ],
        "activities": activities,
    }


# ---------------------------------------------------------------- 先行の解決
def _by_no(tasks: List[Row]) -> Dict[str, Row]:
    by_no: Dict[str, Row] = {}
    for row in tasks:
        by_no.setdefault(str(row.no).strip(), row)
    return by_no


def _valid_predecessors(row: Row, by_no: Dict[str, Row]) -> List[Row]:
    """先行の欄に書かれた項番のうち、この表で見つかったものを解決する。

    見つからない項番は :mod:`wbsgen.daily` の整合性チェックが別に知らせるので、
    図では読み飛ばす。
    """
    found = []
    for no in predecessors(row):
        before = by_no.get(no)
        if before is not None and before.row != row.row:
            found.append(before)
    return found


# ---------------------------------------------------------------- 段・レーン
def _levels(tasks: List[Row], by_no: Dict[str, Row]) -> Dict[int, int]:
    """各行の段 (:func:`_graph_levels` をタスクの前後関係に適用する)。"""
    nodes = [row.row for row in tasks]
    edges = [(before.row, row.row) for row in tasks
             for before in _valid_predecessors(row, by_no)]
    return _graph_levels(nodes, edges)


def _lanes(tasks: List[Row], level: Dict[int, int],
           by_no: Dict[str, Row]) -> Dict[int, int]:
    """段の中での上下位置 (:func:`_graph_lanes` をタスクの前後関係に適用する)。"""
    nodes = [row.row for row in tasks]
    edges = [(before.row, row.row) for row in tasks
             for before in _valid_predecessors(row, by_no)]
    return _graph_lanes(nodes, level, edges)


def _graph_levels(nodes: List[int], edges: List[Edge]) -> Dict[int, int]:
    """ノードの段。循環参照は先行が無かったものとして打ち切る (無限再帰を避ける)。"""
    incoming: Dict[int, List[int]] = {}
    for source, target in edges:
        incoming.setdefault(target, []).append(source)

    level: Dict[int, int] = {}

    def resolve(node: int, visiting: set) -> int:
        if node in level:
            return level[node]
        if node in visiting:
            return 0
        visiting.add(node)
        lvl = 0
        for source in incoming.get(node, []):
            lvl = max(lvl, resolve(source, visiting) + 1)
        visiting.discard(node)
        level[node] = lvl
        return lvl

    for node in nodes:
        resolve(node, set())
    return level


def _graph_lanes(nodes: List[int], level: Dict[int, int],
                  edges: List[Edge]) -> Dict[int, int]:
    """段の中での上下位置。先行 (向かってくる矢印の元) の位置の平均に近い側へ
    寄せて、矢印の交差を減らす (バリセンタ法)。先行の無い段はもともとの
    並び順のままにする。
    """
    order = {node: i for i, node in enumerate(nodes)}
    by_level: Dict[int, List[int]] = {}
    for node in nodes:
        by_level.setdefault(level[node], []).append(node)

    sources_of: Dict[int, List[int]] = {}
    for source, target in edges:
        sources_of.setdefault(target, []).append(source)

    lane: Dict[int, int] = {}
    for lvl in sorted(by_level):
        group = by_level[lvl]
        if lvl == 0:
            ordered = group
        else:
            def barycenter(node: int) -> tuple:
                # まだレーンが決まっていない先行 (循環参照) は無視する
                sources = [lane[s] for s in sources_of.get(node, []) if s in lane]
                average = sum(sources) / len(sources) if sources else float(order[node])
                return (average, order[node])

            ordered = sorted(group, key=barycenter)
        for i, node in enumerate(ordered):
            lane[node] = i
    return lane


# ---------------------------------------------------------------- 最早・最遅
def _times(tasks: List[Row], by_no: Dict[str, Row],
           level: Dict[int, int]) -> Dict[int, Dict[str, int]]:
    """各行の最早開始 (ES)・最早終了 (EF)・最遅開始 (LS)・最遅終了 (LF)・余裕。

    日数 (稼働日) と先行関係だけで数える相対日で、カレンダーの日付とは
    無関係。先行の無い行は ES=0、プロジェクトの終わりは最も遅い EF。
    日数が書かれていない行は 0 日として扱う。段が進む向きの矢印だけを
    使うので、循環参照があっても止まらない。
    """
    duration = {row.row: max(int(row.days or 0), 0) for row in tasks}
    before: Dict[int, List[int]] = {row.row: [] for row in tasks}
    after: Dict[int, List[int]] = {row.row: [] for row in tasks}
    for row in tasks:
        for pred in _valid_predecessors(row, by_no):
            if level[pred.row] < level[row.row]:
                before[row.row].append(pred.row)
                after[pred.row].append(row.row)

    order = sorted(duration, key=lambda key: level[key])
    es: Dict[int, int] = {}
    ef: Dict[int, int] = {}
    for key in order:
        es[key] = max((ef[p] for p in before[key]), default=0)
        ef[key] = es[key] + duration[key]

    end = max(ef.values(), default=0)
    ls: Dict[int, int] = {}
    lf: Dict[int, int] = {}
    for key in reversed(order):
        lf[key] = min((ls[n] for n in after[key]), default=end)
        ls[key] = lf[key] - duration[key]

    return {
        key: {"duration": duration[key], "es": es[key], "ef": ef[key],
              "ls": ls[key], "lf": lf[key], "float": ls[key] - es[key]}
        for key in duration
    }


# ---------------------------------------------------------------- プレジデンス図
def _node(row: Row, level: int, lane: int, time: Dict[str, int]) -> Dict[str, Any]:
    return {
        **time,
        "critical": time["float"] == 0,
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


def _edges(tasks: List[Row], by_no: Dict[str, Row],
           times: Dict[int, Dict[str, int]]) -> List[Dict[str, Any]]:
    """先行のうち、この表に見つかったものだけを矢印にする。

    ``critical`` は、余裕が 0 の同士をつなぎ、後続の開始を先行の終了が
    そのまま決めている矢印 (クリティカルパス上の矢印)。
    """
    return [
        {
            "from": before.row,
            "to": row.row,
            "critical": (times[before.row]["float"] == 0 and times[row.row]["float"] == 0
                         and times[before.row]["ef"] == times[row.row]["es"]),
            # 先行タスクの予定終了日より前に始まる予定になっているか
            "late": bool(before.end and row.start and before.end > row.start),
        }
        for row in tasks
        for before in _valid_predecessors(row, by_no)
    ]


# ---------------------------------------------------------------- アロー図
def _activity(row: Row, tail: int, head: int, by_no: Dict[str, Row]) -> Dict[str, Any]:
    preds = _valid_predecessors(row, by_no)
    return {
        "from": tail,
        "to": head,
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
        # 先行のいずれかの予定終了日より前に始まる予定になっているか
        "late": any(p.end and row.start and p.end > row.start for p in preds),
        "dummy": False,
    }


def _iso(value):
    return value.isoformat() if value else None


__all__ = ["build", "build_arrow"]
