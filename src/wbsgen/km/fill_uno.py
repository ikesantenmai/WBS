"""LibreOffice (UNO) で、同梱の空テンプレートに WBS の中身を書き込む。

システムの python3 (python3-uno が入っているもの) で動かす単独のスクリプト。
``python fill_uno.py テンプレート 出力先 データ.json`` の形で、wbsgen 本体から呼ぶ。
テンプレートの ActiveX ボタンや名前、マクロはそのまま残る。
"""

import datetime
import json
import sys
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

#: シートのコード (VBA) に書かれる、ActiveX ボタンの宣言 ("名前, 図形の番号, 順番, 種類")。
#: LibreOffice はこの宣言を保存時に落とすので、書き足す。Excel はこれでボタンと名前を結び付ける
#: (無いとボタンが出ない)。番号は、元のテンプレートの図形の番号と同じ。
CONTROLS = {
    "Sheet1": ["BTN_SortEx1Disp, 1, 0", "BTN_SortEx2Disp, 2, 1", "BTN_SortAllDisp, 3, 2",
               "BTN_MenuWndDisp, 4, 3", "BTN_TodoAllDisp, 21503, 4", "BTN_TaskWndDisp, 1887, 5"],
    "Sheet2": ["BTN_MemberList_Get, 1, 0", "BTN_MemberList_FilePut, 49, 1",
               "BTN_MemberList_FileGet, 50, 2", "BTN_MemberList_AddGet, 159, 3",
               "BTN_MemberList_FilePutMem, 319, 4"],
    "Sheet3": ["BTN_Config_Rewrite, 8, 7", "BTN_Config_DataShift, 46, 10",
               "BTN_Config_AllDataDel, 714, 11", "BTN_Config_Template1Set, 196, 12",
               "BTN_Config_Template2Set, 197, 13", "BTN_Config_Template1Get, 247, 14",
               "BTN_Config_Template2Get, 248, 15", "BTN_Config_BackPictGet, 307, 16",
               "BTN_Config_Repair, 536, 17"],
    "Sheet4": ["BTN_BaseProcess_FilePut, 1, 0", "BTN_BaseProcess_FileGet, 2, 1"],
}


def prop(name, value):
    item = PropertyValue()
    item.Name, item.Value = name, value
    return item


def serial(text):
    return (datetime.date.fromisoformat(text) - EPOCH).days


def main(template, out, data_file):
    data = json.load(open(data_file, encoding="utf-8"))
    pipe = data["pipe"]          # wbsgen 本体が常駐させている LibreOffice の接続口
    ctx = None
    for _ in range(300):         # 起動直後は、つながるまで待つ
        try:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local)
            ctx = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    if ctx is None:
        raise SystemExit("LibreOffice に接続できません (起動が間に合いませんでした)")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL(
        "file://" + template, "_blank", 0, (prop("Hidden", True),))
    try:
        fill(doc, data)
        declare_controls(doc)
        add_button_macro(doc)
        doc.storeToURL("file://" + out, (prop("FilterName", "MS Excel 97"),))
    finally:
        doc.close(True)


def declare_controls(doc):
    """各シートのコードの先頭に、ActiveX ボタンの宣言 (Attribute VB_Control) を書き足す。"""
    library = doc.BasicLibraries.getByName("VBAProject")
    for module, controls in CONTROLS.items():
        if not library.hasByName(module):
            continue
        source = library.getByName(module)
        lines = "".join(f'Attribute VB_Control = "{c}, MSForms, CommandButton"\n' for c in controls)
        if "VB_Control" not in source:
            library.replaceByName(module, lines + source)


#: 元のテンプレートで、最初は非表示のボタン (抽出ボタン)。残りは表示
HIDDEN = {"BTN_SortEx1Disp", "BTN_SortEx2Disp", "BTN_SortAllDisp"}


def add_button_macro(doc):
    """ブックを開いたとき、足りない ActiveX ボタンを Excel 自身に作らせるマクロを足す。

    ファイルの中のボタンを Excel が読み込めなかったときの備え。同じ名前のボタンが
    既にあれば何もしない。作ったボタンには、シートのコードにある BTN_*_Click が
    名前でつながる。
    """
    sheets = {doc.Sheets.getByName(n).CodeName: doc.Sheets.getByName(n)
              for n in doc.Sheets.ElementNames}
    calls = []
    for module in sorted(CONTROLS):
        sheet = sheets.get(module)
        if sheet is None:
            continue
        page = sheet.DrawPage
        for i in range(page.Count):
            shape = page.getByIndex(i)
            if shape.ShapeType != "com.sun.star.drawing.ControlShape":
                continue
            model = shape.Control
            caption = str(getattr(model, "Label", "")).replace('"', '""')
            point = 72.0 / 2540.0                      # 1/100 mm -> ポイント
            calls.append(
                f'    WBSGEN_AddButton {module}, "{shape.Name}", "{caption}", '
                f'{shape.Position.X * point:.1f}, {shape.Position.Y * point:.1f}, '
                f'{shape.Size.Width * point:.1f}, {shape.Size.Height * point:.1f}, '
                f'{"False" if shape.Name in HIDDEN else "True"}')
    code = (
        "\n'--------------------------------------------------------\n"
        "' wbsgen: ボタンが無ければ作る (ブックを開いたときに呼ぶ)\n"
        "Private Sub WBSGEN_EnsureButtons()\n"
        "    On Error Resume Next\n" + "\n".join(calls) + "\n"
        "End Sub\n\n"
        "Private Sub WBSGEN_AddButton(ws As Object, nm As String, cap As String, _\n"
        "        l As Double, t As Double, w As Double, h As Double, vis As Boolean)\n"
        "    Dim o As Object\n"
        "    On Error Resume Next\n"
        "    Set o = ws.OLEObjects(nm)\n"
        "    If Not o Is Nothing Then Exit Sub\n"
        "    Set o = ws.OLEObjects.Add(ClassType:=\"Forms.CommandButton.1\", _\n"
        "        Left:=l, Top:=t, Width:=w, Height:=h)\n"
        "    If o Is Nothing Then Exit Sub\n"
        "    o.Name = nm\n"
        "    o.Object.Caption = cap\n"
        "    o.Object.TakeFocusOnClick = False\n"
        "    o.Visible = vis\n"
        "End Sub\n")
    library = doc.BasicLibraries.getByName("VBAProject")
    source = library.getByName("ThisWorkbook")
    marker = "Private Sub Workbook_Open()"
    if marker not in source or "WBSGEN_EnsureButtons" in source:
        return
    source = source.replace(marker, marker + "\n    Call WBSGEN_EnsureButtons", 1)
    library.replaceByName("ThisWorkbook", source + code)


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
    # 祝祭日: 設定シートの「休日設定(個別)」(DEFC_Calender_Off = M5:M23、あふれた分は M28:M45)
    slots = [(12, y) for y in range(4, 23)] + [(12, y) for y in range(27, 45)]
    for (x, y), day in zip(slots, data.get("holidays", [])):
        settings.getCellByPosition(x, y).Value = serial(day)
    if data.get("start"):
        settings.getCellRangeByName("C16").Value = serial(data["start"])      # DEFC_Chart_StartDay
    if data.get("period_days"):
        settings.getCellRangeByName("C19").Value = data["period_days"]        # DEFC_Chart_Span
    doc.DocumentProperties.Title = data.get("title", "") or "開発マイルストーン"


if __name__ == "__main__":
    main(*sys.argv[1:4])
