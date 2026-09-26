"""ネットワーク図 (先行関係の依存グラフ) のテスト。"""

import datetime as dt

from wbsgen.importer import Row
from wbsgen.network import build

D = dt.date


def _row(number, no, name="作業", **kwargs):
    return Row(row=number, no=no, name=name, **kwargs)


def _node(model, no):
    return next(n for n in model["nodes"] if n["no"] == no)


# ---------------------------------------------------------------- ノード
def test_a_row_without_a_number_is_not_a_node():
    """大項目だけのまとめ行は、項番が無いので対象にしない。"""
    rows = [_row(2, "", name="開発"), _row(3, "1")]
    model = build(rows)
    assert [n["no"] for n in model["nodes"]] == ["1"]


def test_a_node_carries_the_row_fields():
    rows = [_row(2, "1", name="要件定義", group="開発", member="設計",
                 start=D(2026, 4, 1), end=D(2026, 4, 14), progress=0.5,
                 status="実行中", delay=3)]
    node = _node(build(rows), "1")
    assert node == {
        "row": 2, "no": "1", "name": "要件定義", "group": "開発", "subgroup": "",
        "member": "設計", "start": "2026-04-01", "end": "2026-04-14",
        "progress": 0.5, "status": "実行中", "delay": 3, "level": 0, "lane": 0,
    }


# ---------------------------------------------------------------- 矢印
def test_a_predecessor_becomes_an_edge():
    rows = [
        _row(2, "1", start=D(2026, 4, 1), end=D(2026, 4, 14)),
        _row(3, "2", predecessor="1", start=D(2026, 4, 15), end=D(2026, 4, 20)),
    ]
    model = build(rows)
    assert model["edges"] == [{"from": 2, "to": 3, "late": False}]


def test_several_predecessors_become_several_edges():
    rows = [
        _row(2, "1"),
        _row(3, "2"),
        _row(4, "3", predecessor="1, 2"),
    ]
    edges = {(e["from"], e["to"]) for e in build(rows)["edges"]}
    assert edges == {(2, 4), (3, 4)}


def test_an_unknown_predecessor_is_silently_skipped():
    """先行が見つからない場合は矢印を引かない (:mod:`wbsgen.daily` が別に知らせる)。"""
    rows = [_row(2, "1", predecessor="99")]
    assert build(rows)["edges"] == []


def test_a_task_starting_before_its_predecessor_ends_is_flagged_late():
    rows = [
        _row(2, "1", start=D(2026, 4, 1), end=D(2026, 4, 18)),
        _row(3, "2", predecessor="1", start=D(2026, 4, 10), end=D(2026, 4, 20)),
    ]
    edge = build(rows)["edges"][0]
    assert edge["late"] is True


# ---------------------------------------------------------------- 段
def test_a_task_without_predecessors_is_at_level_zero():
    rows = [_row(2, "1")]
    assert _node(build(rows), "1")["level"] == 0


def test_a_chain_advances_one_level_at_a_time():
    rows = [
        _row(2, "1"),
        _row(3, "2", predecessor="1"),
        _row(4, "3", predecessor="2"),
    ]
    model = build(rows)
    assert [_node(model, no)["level"] for no in ("1", "2", "3")] == [0, 1, 2]


def test_the_level_follows_the_latest_predecessor():
    """先行を 2 つ持つとき、段はより深いほうに 1 足した値になる。"""
    rows = [
        _row(2, "1"),
        _row(3, "2", predecessor="1"),
        _row(4, "3"),
        _row(5, "4", predecessor="2, 3"),
    ]
    model = build(rows)
    assert _node(model, "4")["level"] == 2


def test_independent_chains_keep_their_original_order_at_level_zero():
    rows = [_row(2, "1"), _row(3, "2")]
    model = build(rows)
    assert _node(model, "1")["lane"] == 0
    assert _node(model, "2")["lane"] == 1


def test_lanes_follow_the_predecessors_position():
    """矢印が交差しないよう、先行の並びに合わせて次の段の並びを変える。"""
    rows = [
        _row(2, "1"),                        # 段 0, レーン 0
        _row(3, "2"),                        # 段 0, レーン 1
        _row(4, "3", predecessor="2"),       # 元の並びは先だが、先行 (2) はレーン 1
        _row(5, "4", predecessor="1"),       # 元の並びは後だが、先行 (1) はレーン 0
    ]
    model = build(rows)
    # 先行のレーンに引きずられて、段 1 では並びが入れ替わる
    assert _node(model, "4")["lane"] == 0
    assert _node(model, "3")["lane"] == 1


def test_a_cycle_does_not_hang():
    """循環参照でも、いずれかの矢印を無かったことにして必ず終わる。"""
    rows = [
        _row(2, "1", predecessor="2"),
        _row(3, "2", predecessor="1"),
    ]
    model = build(rows)
    assert {n["level"] for n in model["nodes"]} == {1, 2}
