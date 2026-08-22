"""コマンドラインインタフェース。"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from pathlib import Path

from . import __version__, workbook
from .blank import DEFAULT_ROWS, SpecError, build, parse_date
from .i18n import DEFAULT_LANGUAGE, LANGUAGES, cli as text, labels, normalize
from .timeline import VALID_UNITS, Timeline
from .workcal import WEEKDAY_JP, WEEKDAY_KEYS


def main(argv=None) -> int:
    # ヘルプ自体を訳すので、パーサを組み立てる前に言語を決める必要がある
    lang = _language_from_argv(argv if argv is not None else sys.argv[1:])

    parser = argparse.ArgumentParser(prog="wbsgen", description=text(lang, "description"))
    parser.add_argument("--version", action="version", version=f"wbsgen {__version__}")
    parser.add_argument("--lang", choices=LANGUAGES, default=DEFAULT_LANGUAGE,
                        help=text(lang, "opt_lang"))
    sub = parser.add_subparsers(dest="command", required=True)

    make = sub.add_parser("new", help=text(lang, "new_help"),
                          description=text(lang, "new_description"))
    make.add_argument("-o", "--output", default="wbs.xlsx",
                      help=text(lang, "opt_output", default="wbs.xlsx"))
    make.add_argument("--start", required=True, help=text(lang, "opt_start"))
    period = make.add_mutually_exclusive_group()
    period.add_argument("--end", help=text(lang, "opt_end"))
    period.add_argument("--period-days", type=int, help=text(lang, "opt_period_days"))
    period.add_argument("--months", type=int, help=text(lang, "opt_months"))
    make.add_argument("--unit", choices=VALID_UNITS, default="week",
                      help=text(lang, "opt_unit"))
    make.add_argument("--rows", type=int, default=DEFAULT_ROWS,
                      help=text(lang, "opt_rows", default=DEFAULT_ROWS))
    make.add_argument("--title", default=None, help=text(lang, "opt_title"))
    make.add_argument("--member", action="append", default=None,
                      help=text(lang, "opt_member"))
    make.add_argument("--workdays", default=None, help=text(
        lang, "opt_workdays",
        keys=f"{','.join(WEEKDAY_KEYS)} / {'・'.join(WEEKDAY_JP)}"))
    make.add_argument("--holiday", action="append", default=None,
                      help=text(lang, "opt_holiday"))
    make.add_argument("--no-jp-holidays", action="store_true", help=text(lang, "opt_no_jp"))
    make.add_argument("--lang", choices=LANGUAGES, default=DEFAULT_LANGUAGE,
                      help=text(lang, "opt_lang"))
    make.set_defaults(func=_cmd_new)

    host, port, hosted = _serve_defaults()
    serve = sub.add_parser("serve", help=text(lang, "serve_help"),
                           description=text(lang, "serve_description"))
    serve.add_argument("--host", default=host, help=text(lang, "opt_host", default=host))
    serve.add_argument("--port", type=int, default=port,
                       help=text(lang, "opt_port", default=port))
    serve.add_argument("--reload", action="store_true", help=text(lang, "opt_reload"))
    serve.add_argument("--lang", choices=LANGUAGES, default=DEFAULT_LANGUAGE,
                       help=text(lang, "opt_lang"))
    serve.set_defaults(func=_cmd_serve, hosted=hosted)

    args = parser.parse_args(argv)
    args.language = normalize(getattr(args, "lang", None) or lang)
    try:
        return args.func(args)
    except SpecError as exc:
        print(text(args.language, "error", reason=exc), file=sys.stderr)
        return 2


def _language_from_argv(argv) -> str:
    """``--lang`` をパーサの前に拾う (ヘルプの文言を決めるため)。"""
    for index, item in enumerate(argv):
        if item == "--lang" and index + 1 < len(argv):
            return normalize(argv[index + 1])
        if item.startswith("--lang="):
            return normalize(item.split("=", 1)[1])
    return DEFAULT_LANGUAGE


# ----------------------------------------------------------------------
def _cmd_new(args) -> int:
    lang = args.language
    spec = build(
        start=parse_date(args.start, "--start", lang),
        end=parse_date(args.end, "--end", lang) if args.end else None,
        period_days=args.period_days,
        months=args.months,
        unit=args.unit,
        rows=args.rows,
        title=args.title,
        members=args.member,
        workdays=_workdays(args.workdays),
        japanese_holidays=not args.no_jp_holidays,
        holidays=[parse_date(d, "--holiday", lang) for d in (args.holiday or [])],
        language=lang,
    )
    output = Path(args.output)
    workbook.write(spec, output)

    columns = len(Timeline(spec.start, spec.period_days, spec.unit,
                           spec.calendar(), spec.language))
    print(text(lang, "done", path=output))
    print(text(lang, "out_period", start=spec.start, end=spec.end, days=spec.period_days))
    print(text(lang, "out_chart", unit=args.unit, columns=columns))
    print(text(lang, "out_rows", rows=spec.rows))
    if spec.members:
        print(text(lang, "out_members", members=", ".join(spec.members)))
    return 0


def _workdays(value):
    if not value:
        return None
    return [part.strip() for part in str(value).replace("、", ",").split(",") if part.strip()]


# ----------------------------------------------------------------------
def _serve_defaults():
    """待ち受け先の既定値を決める。

    PaaS (Render / Heroku / Cloud Run など) は待ち受けポートを環境変数 ``PORT``
    で渡してくる。その場合は ``0.0.0.0`` で待ち受けないと外部から届かないので、
    ホストの既定も切り替える。手元で動かすときは従来どおり ``127.0.0.1``。
    """
    raw = os.environ.get("PORT")
    hosted = bool(raw)
    try:
        port = int(raw) if raw else 8000
    except ValueError:
        port = 8000
        hosted = False
    host = os.environ.get("HOST") or ("0.0.0.0" if hosted else "127.0.0.1")
    return host, port, hosted


def _cmd_serve(args) -> int:
    lang = args.language
    try:
        from .web import serve
    except ImportError:
        print(text(lang, "need_web"), file=sys.stderr)
        return 2

    if getattr(args, "hosted", False) and args.host == "0.0.0.0":
        print(text(lang, "serving_hosted", port=args.port))
    else:
        print(text(lang, "serving", host=args.host, port=args.port))
    serve(host=args.host, port=args.port, reload=args.reload)
    return 0


if __name__ == "__main__":
    sys.exit(main())
