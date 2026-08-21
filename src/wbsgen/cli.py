"""コマンドラインインタフェース。"""

from __future__ import annotations

import argparse
import datetime as _dt
import sys
from pathlib import Path

from . import __version__, workbook
from .blank import DEFAULT_ROWS, SpecError, build, parse_date
from .timeline import VALID_UNITS
from .workcal import WEEKDAY_JP, WEEKDAY_KEYS


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="wbsgen",
        description="期間を指定して、中身が空の WBS (ガントチャート用紙) を作ります。",
    )
    parser.add_argument("--version", action="version", version=f"wbsgen {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    make = sub.add_parser(
        "new",
        help="空の WBS を作る",
        description="日程表と記入用の空行だけの Excel を書き出します。",
    )
    make.add_argument("-o", "--output", default="wbs.xlsx", help="出力先 (既定: wbs.xlsx)")
    make.add_argument("--start", required=True, help="開始日 (YYYY-MM-DD)")
    period = make.add_mutually_exclusive_group()
    period.add_argument("--end", help="終了日 (YYYY-MM-DD)")
    period.add_argument("--period-days", type=int, help="期間を暦日で指定する")
    period.add_argument("--months", type=int, help="期間を月数で指定する (既定: 12)")
    make.add_argument("--unit", choices=VALID_UNITS, default="week",
                      help="日程表の単位 (既定: week)")
    make.add_argument("--rows", type=int, default=DEFAULT_ROWS,
                      help=f"記入用の空行数 (既定: {DEFAULT_ROWS})")
    make.add_argument("--title", default=None, help="プロジェクト名")
    make.add_argument("--member", action="append", default=None,
                      help="担当者を担当者一覧に載せる (複数指定可)")
    make.add_argument("--workdays", default=None,
                      help="稼働曜日をカンマ区切りで指定する "
                           f"({','.join(WEEKDAY_KEYS)} / {'・'.join(WEEKDAY_JP)}。"
                           "既定: mon,tue,wed,thu,fri)")
    make.add_argument("--holiday", action="append", default=None,
                      help="休業日を追加する (YYYY-MM-DD、複数指定可)")
    make.add_argument("--no-jp-holidays", action="store_true",
                      help="日本の祝日を休日として扱わない")
    make.set_defaults(func=_cmd_new)

    serve = sub.add_parser("serve", help="Web アプリケーションを起動する")
    serve.add_argument("--host", default="127.0.0.1", help="待ち受けホスト (既定: 127.0.0.1)")
    serve.add_argument("--port", type=int, default=8000, help="待ち受けポート (既定: 8000)")
    serve.add_argument("--reload", action="store_true", help="コード変更時に自動再起動する")
    serve.set_defaults(func=_cmd_serve)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except SpecError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 2


# ----------------------------------------------------------------------
def _cmd_new(args) -> int:
    spec = build(
        start=parse_date(args.start, "--start"),
        end=parse_date(args.end, "--end") if args.end else None,
        period_days=args.period_days,
        months=args.months,
        unit=args.unit,
        rows=args.rows,
        title=args.title,
        members=args.member,
        workdays=_workdays(args.workdays),
        japanese_holidays=not args.no_jp_holidays,
        holidays=[parse_date(d, "--holiday") for d in (args.holiday or [])],
    )
    output = Path(args.output)
    workbook.write(spec, output)

    print(f"生成しました: {output}")
    print(f"  期間        : {spec.start} 〜 {spec.end} ({spec.period_days} 日)")
    print(f"  日程表      : {args.unit} / {len(_columns(spec))} 列")
    print(f"  記入用の空行: {spec.rows} 行")
    if spec.members:
        print(f"  担当者      : {', '.join(spec.members)}")
    return 0


def _columns(spec):
    from .timeline import Timeline

    return Timeline(spec.start, spec.period_days, spec.unit, spec.calendar()).columns


def _workdays(text):
    if not text:
        return None
    return [part.strip() for part in str(text).replace("、", ",").split(",") if part.strip()]


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


if __name__ == "__main__":
    sys.exit(main())
