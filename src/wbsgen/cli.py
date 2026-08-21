"""コマンドラインインタフェース。"""

from __future__ import annotations

import argparse
import datetime as _dt
import sys
from pathlib import Path

from . import __version__
from .loader import ProjectError, load
from .render import WorkbookRenderer
from .scheduling import resolve_project
from .templates import TEMPLATE_NAMES, write_template


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="wbsgen",
        description="WBS / ガントチャート Excel を生成します。",
    )
    parser.add_argument("--version", action="version", version=f"wbsgen {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    build = sub.add_parser("build", help="定義ファイルから Excel を生成する")
    build.add_argument("source", help="プロジェクト定義 (.yaml / .json / .csv)")
    build.add_argument("-o", "--output", default=None, help="出力先 .xlsx (既定: <定義名>.xlsx)")
    build.add_argument("--base-date", default=None,
                       help="現在日 (YYYY-MM-DD)。状態・遅れの判定基準日")
    build.add_argument("--unit", choices=("day", "week", "month"), default=None,
                       help="チャート表示単位を上書きする")
    build.set_defaults(func=_cmd_build)

    init = sub.add_parser("init", help="定義ファイルの雛形を書き出す")
    init.add_argument("-o", "--output", default="wbs.yaml", help="出力先 (既定: wbs.yaml)")
    init.add_argument("-t", "--template", choices=TEMPLATE_NAMES, default="standard",
                      help="雛形の種類 (既定: standard)")
    init.add_argument("-f", "--force", action="store_true", help="既存ファイルを上書きする")
    init.set_defaults(func=_cmd_init)

    check = sub.add_parser("check", help="定義ファイルを検証し、日程の要約を表示する")
    check.add_argument("source", help="プロジェクト定義")
    check.set_defaults(func=_cmd_check)

    serve = sub.add_parser("serve", help="Web アプリケーションを起動する")
    serve.add_argument("--host", default="127.0.0.1", help="待ち受けホスト (既定: 127.0.0.1)")
    serve.add_argument("--port", type=int, default=8000, help="待ち受けポート (既定: 8000)")
    serve.add_argument("--reload", action="store_true", help="コード変更時に自動再起動する")
    serve.set_defaults(func=_cmd_serve)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except ProjectError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 2
    except FileNotFoundError as exc:
        print(f"エラー: ファイルが見つかりません: {exc.filename}", file=sys.stderr)
        return 2


# ----------------------------------------------------------------------
def _cmd_build(args) -> int:
    source = Path(args.source)
    project = load(source)
    if args.base_date:
        project.chart.base_date = _dt.datetime.strptime(args.base_date, "%Y-%m-%d").date()
    if args.unit:
        project.chart.unit = args.unit

    output = Path(args.output) if args.output else source.with_suffix(".xlsx")
    renderer = WorkbookRenderer(project)
    renderer.save(output)

    print(f"生成しました: {output}")
    print(f"  タスク数    : {len(project.tasks)}")
    print(f"  表示期間    : {renderer.timeline.start} 〜 {renderer.timeline.end}"
          f" ({len(renderer.timeline)} 列 / {project.chart.unit})")
    print(f"  基準日      : {project.chart.base_date or _dt.date.today()}")
    return 0


def _cmd_init(args) -> int:
    output = Path(args.output)
    if output.exists() and not args.force:
        print(f"エラー: すでに存在します: {output} (--force で上書き)", file=sys.stderr)
        return 2
    write_template(args.template, output)
    print(f"雛形を書き出しました: {output}")
    print(f"  wbsgen build {output} で Excel を生成できます。")
    return 0


def _cmd_serve(args) -> int:
    try:
        from .web import serve
    except ImportError:
        print("エラー: Web アプリには追加の依存が必要です。"
              "\n  pip install 'wbsgen[web]'", file=sys.stderr)
        return 2
    print(f"起動しました: http://{args.host}:{args.port}/  (Ctrl+C で終了)")
    serve(host=args.host, port=args.port, reload=args.reload)
    return 0


def _cmd_check(args) -> int:
    project = load(Path(args.source))
    resolve_project(project)
    base = project.chart.base_date or _dt.date.today()

    delayed = [t for t in project.tasks if t.delay]
    done = [t for t in project.tasks if t.is_complete]
    running = [t for t in project.tasks if t.actual_start and not t.is_complete]

    print(f"プロジェクト: {project.title}")
    print(f"基準日      : {base}")
    print(f"タスク      : 全 {len(project.tasks)} 件"
          f" / 完了 {len(done)} / 進行中 {len(running)} / 遅延 {len(delayed)}")
    ends = [t.end for t in project.tasks if t.end]
    if ends:
        print(f"最終予定日  : {max(ends)}")
    if delayed:
        print("\n遅延タスク:")
        width = max(len(t.name) for t in delayed)
        for task in sorted(delayed, key=lambda t: -(t.delay or 0)):
            print(f"  {task.name:<{width}}  {task.status}  (予定終了 {task.end})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
