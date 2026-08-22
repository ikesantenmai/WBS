"""日本語版・英語版の文言。

Excel に書き出すシート名・見出し・表示形式と、利用者に見せるメッセージを
言語ごとにまとめる。既定は日本語。

読み込み側は言語を問わず両方の見出しを受け付けるので、日本語で作った
ファイルを英語表示で開く (またはその逆) こともできる。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Tuple

DEFAULT_LANGUAGE = "ja"
LANGUAGES = ("ja", "en")


def normalize(language) -> str:
    """``ja-JP`` や ``EN`` のような指定を ``ja`` / ``en`` に丸める。"""
    if not language:
        return DEFAULT_LANGUAGE
    key = str(language).strip().lower().replace("_", "-").split("-")[0]
    return key if key in LANGUAGES else DEFAULT_LANGUAGE


@dataclass(frozen=True)
class Labels:
    """Excel に書き出す文言と表示形式。"""

    sheet_plan: str
    sheet_member: str
    sheet_config: str

    #: 表側の列見出し (属性名 -> 見出し)
    columns: Dict[str, str]
    #: 予定 / 実績 のグループ見出し
    group_plan: str
    group_actual: str

    #: 曜日 (月曜始まり)
    weekdays: Tuple[str, ...]
    #: 月の呼び方 (1 月から)
    month_names: Tuple[str, ...]
    #: 年の書き方
    year_format: str
    #: 表示単位の名前
    units: Dict[str, str]
    #: 表示単位ごとの (上段, 下段) の数値書式
    header_formats: Dict[str, Tuple[str, str]]
    #: 記入欄の数値書式
    date_format: str
    days_format: str

    #: 担当者一覧シート
    member_title: str
    member_headers: Tuple[str, str, str]

    #: 設定シート
    config_chart: str
    config_start: str
    config_end: str
    config_period: str
    config_unit: str
    config_rows: str
    config_workdays: str
    config_holidays: str
    weekday_names: Tuple[str, ...]
    work_on: str
    work_off: str

    #: 状態の文言 (``{days}`` に日数が入る)
    status_done: str
    status_delayed: str
    status_remaining: str
    status_upcoming: str
    status_none: str

    #: 既定のプロジェクト名 (``{year}`` に開始年が入る)
    default_title: str


JA = Labels(
    sheet_plan="スケジュール",
    sheet_member="担当者一覧",
    sheet_config="設定",
    columns={
        "group": "大項目", "subgroup": "中項目", "no": "項番", "name": "項目",
        "start": "開始日", "days": "日数", "end": "終了日",
        "actual_start": "開始", "actual_days": "日数", "actual_end": "終了",
        "delay": "遅れ", "progress": "進捗",
        "effort": "工数", "predecessor": "先行", "member": "担当", "status": "状態",
    },
    group_plan="予定",
    group_actual="実績",
    weekdays=("月", "火", "水", "木", "金", "土", "日"),
    month_names=tuple(f"{m}月" for m in range(1, 13)),
    year_format="{year}年",
    units={"day": "日単位", "week": "週単位", "month": "月単位"},
    header_formats={
        "day": ('m"月"', "d"),
        "week": ('m"月"', "m/d"),
        "month": ('yyyy"年"', 'm"月"'),
    },
    date_format="m/dd",
    days_format='0\\ "日"',
    member_title="◆担当者一覧",
    member_headers=("担当", "色", "備考"),
    config_chart="◆チャート表示設定",
    config_start="チャート表示開始日",
    config_end="チャート表示終了日",
    config_period="チャート表示期間(日)",
    config_unit="チャート表示単位",
    config_rows="記入用の空行",
    config_workdays="◆作業日設定",
    config_holidays="◆休日一覧",
    weekday_names=("月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"),
    work_on="出",
    work_off="休",
    status_done="完了",
    status_delayed="遅れ {days} 日",
    status_remaining="残り {days} 日",
    status_upcoming="あと {days} 日",
    status_none="-",
    default_title="{year}年 スケジュール",
)

EN = Labels(
    sheet_plan="Schedule",
    sheet_member="Members",
    sheet_config="Settings",
    columns={
        "group": "Group", "subgroup": "Sub-group", "no": "No.", "name": "Task",
        "start": "Start", "days": "Days", "end": "End",
        "actual_start": "Start", "actual_days": "Days", "actual_end": "End",
        "delay": "Delay", "progress": "Progress",
        "effort": "Effort", "predecessor": "Pred.", "member": "Owner", "status": "Status",
    },
    group_plan="Planned",
    group_actual="Actual",
    weekdays=("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"),
    month_names=("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
    year_format="{year}",
    units={"day": "Daily", "week": "Weekly", "month": "Monthly"},
    header_formats={
        "day": ("mmm", "d"),
        "week": ("mmm", "m/d"),
        "month": ("yyyy", "mmm"),
    },
    date_format="m/dd",
    days_format='0\\ "d"',
    member_title="◆Members",
    member_headers=("Owner", "Colour", "Note"),
    config_chart="◆Chart",
    config_start="Chart start",
    config_end="Chart end",
    config_period="Period (days)",
    config_unit="Chart unit",
    config_rows="Blank rows",
    config_workdays="◆Working days",
    config_holidays="◆Holidays",
    weekday_names=("Monday", "Tuesday", "Wednesday", "Thursday",
                   "Friday", "Saturday", "Sunday"),
    work_on="Work",
    work_off="Off",
    status_done="Done",
    status_delayed="Delayed {days} d",
    status_remaining="Remaining {days} d",
    status_upcoming="Starts in {days} d",
    status_none="-",
    default_title="{year} Schedule",
)

LABELS: Dict[str, Labels] = {"ja": JA, "en": EN}


def labels(language=None) -> Labels:
    return LABELS[normalize(language)]


# ----------------------------------------------------------------------
# 状態表示の配色。書かれる言葉は利用者が決めるので、両方の言語の言い回しを
# まとめて見る (日本語で書いたファイルを英語表示で開いても色が付く)。
# ----------------------------------------------------------------------
#: 先頭一致で見る (状態の種類 -> その言い回し)
STATUS_KEYWORDS = {
    "done": ("完了", "Done", "Complete", "Finished"),
    "running": ("実行中", "In progress", "Running", "Active"),
    "remaining": ("残り", "Remaining", "Left"),
    "delayed": ("遅れ", "Delay", "Delayed", "Late", "Overdue"),
    "upcoming": ("あと", "Starts in", "Upcoming", "Soon"),
}


def status_kind(status: str):
    """状態の文字列から種類を返す。当てはまらなければ ``None``。"""
    text = (status or "").strip()
    if not text:
        return None
    lowered = text.lower()
    for kind, words in STATUS_KEYWORDS.items():
        for word in words:
            if lowered.startswith(word.lower()):
                return kind
    return None


# ----------------------------------------------------------------------
# 利用者に見せるメッセージ
# ----------------------------------------------------------------------
MESSAGES = {
    "ja": {
        "need_start": "開始日を指定してください。",
        "bad_unit": "表示単位が不正です: {value!r} ({choices})",
        "end_before_start": "終了日が開始日より前です: {start} 〜 {end}",
        "period_too_short": "期間は 1 日以上で指定してください: {days}",
        "period_too_long": "期間が長すぎます (上限 10 年): {days} 日",
        "bad_rows": "行数は 0〜{maximum} の範囲で指定してください: {value}",
        "too_many_members": "担当者は {maximum} 人までです: {value} 人",
        "bad_weekday": "曜日の指定が不正です: {value!r}",
        "not_mapping": "指定はマッピングである必要があります。",
        "need_date": "{field} を指定してください (YYYY-MM-DD)。",
        "bad_date": "{field} は YYYY-MM-DD 形式で指定してください: {value!r}",
        "bad_int": "{field} は整数で指定してください: {value!r}",
        "bad_language": "言語の指定が不正です: {value!r} ({choices})",

        "not_excel": "Excel として読めません: {reason}",
        "no_header": "見出し行が見つかりません (「項目」の列が必要です)。",
        "no_task_column": "「項目」の列が見つかりません。"
                          "このツールが作った WBS を記入したファイルを指定してください。",
        "cell_bad_date": "日付として読めません: {value!r}",
        "cell_bad_days": "日数として読めません: {value!r}",
        "cell_bad_number": "数値として読めません: {value!r}",
        "cell_bad_progress": "進捗として読めません: {value!r}",
        "cell_prefix": "{row} 行目「{column}」: {reason}",
        "too_many_rows": "{maximum} 行を超えたので、以降は読み飛ばしました。",
        "imported_title": "読み込んだ WBS",

        "upload_too_large": "ファイルが大きすぎます (上限 {limit}MB)",
        "unsupported_format": "対応していない形式です: {suffix} — .xlsx を指定してください",
        "no_suffix": "(拡張子なし)",
        "no_rows": "記入された行が見つかりません。"
                   "項目と日付を入れてから読み込んでください。",
    },
    "en": {
        "need_start": "Please give a start date.",
        "bad_unit": "Invalid chart unit: {value!r} ({choices})",
        "end_before_start": "The end date is before the start date: {start} - {end}",
        "period_too_short": "The period must be at least one day: {days}",
        "period_too_long": "The period is too long (10 years maximum): {days} days",
        "bad_rows": "Rows must be between 0 and {maximum}: {value}",
        "too_many_members": "At most {maximum} owners are allowed: {value}",
        "bad_weekday": "Invalid weekday: {value!r}",
        "not_mapping": "The request must be a mapping.",
        "need_date": "Please give {field} (YYYY-MM-DD).",
        "bad_date": "{field} must be in YYYY-MM-DD format: {value!r}",
        "bad_int": "{field} must be a whole number: {value!r}",
        "bad_language": "Invalid language: {value!r} ({choices})",

        "not_excel": "Cannot read this as Excel: {reason}",
        "no_header": "No header row found (a \"Task\" column is required).",
        "no_task_column": "No \"Task\" column found."
                          " Please choose a WBS created by this tool and filled in.",
        "cell_bad_date": "Cannot read as a date: {value!r}",
        "cell_bad_days": "Cannot read as a number of days: {value!r}",
        "cell_bad_number": "Cannot read as a number: {value!r}",
        "cell_bad_progress": "Cannot read as progress: {value!r}",
        "cell_prefix": "Row {row}, \"{column}\": {reason}",
        "too_many_rows": "More than {maximum} rows; the rest were skipped.",
        "imported_title": "Imported WBS",

        "upload_too_large": "The file is too large ({limit}MB maximum)",
        "unsupported_format": "Unsupported format: {suffix} - please choose an .xlsx file",
        "no_suffix": "(no extension)",
        "no_rows": "No filled-in rows found."
                   " Please enter tasks and dates before importing.",
    },
}


#: CLI の文言
CLI = {
    "ja": {
        "description": "期間を指定して、中身が空の WBS (ガントチャート用紙) を作ります。",
        "new_help": "空の WBS を作る",
        "new_description": "日程表と記入用の空行だけの Excel を書き出します。",
        "serve_help": "Web アプリケーションを起動する",
        "serve_description": "環境変数 PORT があればそれを使い、外部から届くように "
                             "0.0.0.0 で待ち受けます (Render などの PaaS 向け)。"
                             "そのため、そうした環境では引数なしの `wbsgen serve` "
                             "だけで動きます。",
        "opt_output": "出力先 (既定: {default})",
        "opt_start": "開始日 (YYYY-MM-DD)",
        "opt_end": "終了日 (YYYY-MM-DD)",
        "opt_period_days": "期間を暦日で指定する",
        "opt_months": "期間を月数で指定する (既定: 12)",
        "opt_unit": "日程表の単位 (既定: week)",
        "opt_rows": "記入用の空行数 (既定: {default})",
        "opt_title": "プロジェクト名",
        "opt_member": "担当者を担当者一覧に載せる (複数指定可)",
        "opt_workdays": "稼働曜日をカンマ区切りで指定する ({keys}。既定: mon,tue,wed,thu,fri)",
        "opt_holiday": "休業日を追加する (YYYY-MM-DD、複数指定可)",
        "opt_no_jp": "日本の祝日を休日として扱わない",
        "opt_lang": "表示言語 (既定: ja)",
        "opt_host": "待ち受けホスト (既定: {default}。環境変数 HOST でも指定できます)",
        "opt_port": "待ち受けポート (既定: {default}。環境変数 PORT でも指定できます)",
        "opt_reload": "コード変更時に自動再起動する",
        "done": "生成しました: {path}",
        "out_period": "  期間        : {start} 〜 {end} ({days} 日)",
        "out_chart": "  日程表      : {unit} / {columns} 列",
        "out_rows": "  記入用の空行: {rows} 行",
        "out_members": "  担当者      : {members}",
        "error": "エラー: {reason}",
        "need_web": "エラー: Web アプリには追加の依存が必要です。"
                    "\n  pip install 'wbsgen[web]'",
        "serving": "起動しました: http://{host}:{port}/  (Ctrl+C で終了)",
        "serving_hosted": "起動しました: ポート {port} で待ち受けます "
                          "(環境変数 PORT を検出。Ctrl+C で終了)",
    },
    "en": {
        "description": "Creates an empty WBS (Gantt chart sheet) for a given period.",
        "new_help": "create an empty WBS",
        "new_description": "Writes an Excel file with the calendar header and blank rows.",
        "serve_help": "start the web application",
        "serve_description": "Uses the PORT environment variable when set and listens on "
                             "0.0.0.0 so it is reachable from outside (for Render and "
                             "other PaaS). On those platforms `wbsgen serve` alone is "
                             "enough.",
        "opt_output": "output path (default: {default})",
        "opt_start": "start date (YYYY-MM-DD)",
        "opt_end": "end date (YYYY-MM-DD)",
        "opt_period_days": "period in calendar days",
        "opt_months": "period in months (default: 12)",
        "opt_unit": "calendar unit (default: week)",
        "opt_rows": "number of blank rows (default: {default})",
        "opt_title": "project name",
        "opt_member": "add an owner to the members sheet (repeatable)",
        "opt_workdays": "working weekdays, comma separated ({keys}."
                        " default: mon,tue,wed,thu,fri)",
        "opt_holiday": "add a non-working day (YYYY-MM-DD, repeatable)",
        "opt_no_jp": "do not treat Japanese public holidays as non-working days",
        "opt_lang": "language (default: ja)",
        "opt_host": "host to bind (default: {default}; the HOST variable also works)",
        "opt_port": "port to bind (default: {default}; the PORT variable also works)",
        "opt_reload": "restart automatically when the code changes",
        "done": "Created: {path}",
        "out_period": "  Period : {start} - {end} ({days} days)",
        "out_chart": "  Calendar: {unit} / {columns} columns",
        "out_rows": "  Blank rows: {rows}",
        "out_members": "  Owners : {members}",
        "error": "Error: {reason}",
        "need_web": "Error: the web application needs extra dependencies."
                    "\n  pip install 'wbsgen[web]'",
        "serving": "Started: http://{host}:{port}/  (Ctrl+C to stop)",
        "serving_hosted": "Started: listening on port {port} "
                          "(PORT variable detected. Ctrl+C to stop)",
    },
}


def cli(language, key: str, **values) -> str:
    """CLI の文言を ``language`` で返す。"""
    table = CLI.get(normalize(language), CLI[DEFAULT_LANGUAGE])
    return table[key].format(**values)


def message(language, key: str, **values) -> str:
    """``key`` の文言を ``language`` で返す。"""
    table = MESSAGES.get(normalize(language), MESSAGES[DEFAULT_LANGUAGE])
    return table[key].format(**values)
