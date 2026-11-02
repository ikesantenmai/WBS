"""LibreOffice (UNO) で、同梱の空テンプレートに WBS の中身を書き込む。

システムの python3 (python3-uno が入っているもの) で動かす単独のスクリプト。
``python fill_uno.py テンプレート 出力先 データ.json`` の形で、wbsgen 本体から呼ぶ。
テンプレートの ActiveX ボタンや名前、マクロはそのまま残る。
"""

import datetime
import json
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue

EPOCH = datetime.date(1899, 12, 30)
FIRST_ROW = 4          # 0 始まり (= 5 行目)
LAST_ROW = 201         # 202 行目まで
#: 行の項目 -> 列 (0 始まり)。B..R のうち P 列は使わない
COLUMNS = [("group", 1), ("subgroup", 2), ("no", 3), ("name", 4), ("start", 5),
           ("days", 6), ("end", 7), ("actual_start", 8), ("actual_days", 9),
           ("actual_end", 10), ("delay", 11), ("progress", 12), ("effort", 13),
           ("predecessor", 14), ("member", 16), ("status", 17)]
DATES = {"start", "end", "actual_start", "actual_end"}


def prop(name, value):
    item = PropertyValue()
    item.Name, item.Value = name, value
    return item


def serial(text):
    return (datetime.date.fromisoformat(text) - EPOCH).days


def main(template, out, data_file):
    data = json.load(open(data_file, encoding="utf-8"))
    profile = tempfile.mkdtemp(prefix="km_profile_")
    port = data.get("port", 2099)
    proc = subprocess.Popen(
        [data["soffice"], "--headless", "--norestore", "--nologo",
         f"--accept=socket,host=localhost,port={port};urp;",
         f"-env:UserInstallation=file://{profile}"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = None
        for _ in range(90):
            try:
                local = uno.getComponentContext()
                resolver = local.ServiceManager.createInstanceWithContext(
                    "com.sun.star.bridge.UnoUrlResolver", local)
                ctx = resolver.resolve(
                    f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(1)
        if ctx is None:
            raise SystemExit("LibreOffice に接続できません")
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL(
            "file://" + template, "_blank", 0, (prop("Hidden", True),))
        try:
            fill(doc, data)
            doc.storeToURL("file://" + out, (prop("FilterName", "MS Excel 97"),))
        finally:
            doc.close(True)
    finally:
        proc.terminate()
        try:
            proc.wait(10)
        except Exception:
            proc.kill()


def fill(doc, data):
    sheets = doc.Sheets
    schedule = sheets.getByName("スケジュール")
    schedule.getCellByPosition(1, 0).String = data.get("title", "")
    for offset, row in enumerate(data.get("rows", [])):
        y = FIRST_ROW + offset
        if y > LAST_ROW:
            break
        for key, x in COLUMNS:
            value = row.get(key)
            if value in (None, ""):
                continue
            cell = schedule.getCellByPosition(x, y)
            if key in DATES:
                cell.Value = serial(value)
            elif isinstance(value, (int, float)):
                cell.Value = value
            else:
                cell.String = str(value)

    members = sheets.getByName("担当者一覧")
    for offset, name in enumerate(data.get("members", [])[:5]):
        members.getCellByPosition(2, 3 + offset).String = name
        members.getCellByPosition(3, 3 + offset).String = "○"

    settings = sheets.getByName("設定")
    if data.get("start"):
        settings.getCellRangeByName("C16").Value = serial(data["start"])      # DEFC_Chart_StartDay
    if data.get("period_days"):
        settings.getCellRangeByName("C19").Value = data["period_days"]        # DEFC_Chart_Span
    doc.DocumentProperties.Title = data.get("title", "") or "開発マイルストーン"


if __name__ == "__main__":
    main(*sys.argv[1:4])
