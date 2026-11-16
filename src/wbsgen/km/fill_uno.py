"""LibreOffice (UNO) で、同梱の空テンプレートに WBS の中身を書き込む。

システムの python3 (python3-uno が入っているもの) で動かす単独のスクリプト。
``python fill_uno.py テンプレート 出力先 データ.json`` の形で、wbsgen 本体から呼ぶ。
テンプレートの ActiveX ボタンや名前、マクロはそのまま残る。
"""

import datetime
import json
import os
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

#: ボタンのあるシートのコード名 (ボタンはマクロで作る。add_button_macro)
CONTROL_MODULES = ["Sheet1", "Sheet2", "Sheet3", "Sheet4"]


def prop(name, value):
    item = PropertyValue()
    item.Name, item.Value = name, value
    return item


def serial(text):
    return (datetime.date.fromisoformat(text) - EPOCH).days


def main(template, out, data_file, vba_dir):
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
        restore_vba(doc, vba_dir)
        add_button_macro(doc)
        remove_controls(doc)
        doc.storeToURL("file://" + out, (prop("FilterName", "MS Excel 97"),))
    finally:
        doc.close(True)


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
    for module in CONTROL_MODULES:
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
                f'    msg = msg & WBSGEN_AddButton({module}, "{shape.Name}", "{caption}", '
                f'{shape.Position.X * point:.1f}, {shape.Position.Y * point:.1f}, '
                f'{shape.Size.Width * point:.1f}, {shape.Size.Height * point:.1f}, '
                f'{"False" if shape.Name in HIDDEN else "True"})')
    code = (
        "\n'--------------------------------------------------------\n"
        "' wbsgen: ボタンが無ければ作る。ブックの読み込みが終わってから呼ぶ (OnTime)。\n"
        "' 作れなかったときは理由を表示する\n"
        "Public Sub WBSGEN_EnsureButtons()\n"
        "    Dim msg As String, n As Integer, keep As Object, r As String\n"
        "    On Error Resume Next\n"
        "    Set keep = ActiveSheet\n"
        "    Application.ScreenUpdating = False\n" + "\n".join(calls) + "\n"
        "    keep.Activate\n"
        "    Application.ScreenUpdating = True\n"
        "    On Error GoTo 0\n"
        "    If Len(msg) > 0 Then\n"
        "        n = UBound(Split(msg, vbLf))\n"
        "        MsgBox \"ボタンを \" & n & \" 個作れませんでした (1 回だけ表示します):\" & vbLf & Left(msg, 300), vbExclamation, \"wbsgen\"\n"
        "    End If\n"
        "End Sub\n\n"
        "Private Function WBSGEN_AddButton(ws As Object, nm As String, cap As String, _\n"
        "        l As Double, t As Double, w As Double, h As Double, vis As Boolean) As String\n"
        "    Dim o As Object, i As Integer\n"
        "    On Error Resume Next\n"
        "    Set o = ws.OLEObjects(nm)\n"
        "    Err.Clear\n"
        "    If Not o Is Nothing Then Exit Function\n"
        "    ws.Activate                         ' 表示中のシートでないと挿入できないことがある\n"
        "    For i = 1 To 2\n"
        "        Err.Clear\n"
        "        Set o = Nothing\n"
        "        Set o = ws.OLEObjects.Add(ClassType:=\"Forms.CommandButton.1\", _\n"
        "            Left:=l, Top:=t, Width:=w, Height:=h)\n"
        "        If Err.Number = 0 And Not o Is Nothing Then Exit For\n"
        "        DoEvents\n"
        "    Next i\n"
        "    If o Is Nothing Then\n"
        "        WBSGEN_AddButton = ws.Name & \" / \" & nm & \": \" & Err.Number & \" \" & Err.Description & vbLf\n"
        "        Exit Function\n"
        "    End If\n"
        "    Err.Clear\n"
        "    o.Name = nm\n"
        "    o.Object.Caption = cap\n"
        "    o.Object.TakeFocusOnClick = False\n"
        "    o.Visible = vis\n"
        "End Function\n")
    library = doc.BasicLibraries.getByName("VBAProject")
    source = library.getByName("ThisWorkbook")
    marker = "Private Sub Workbook_Open()"
    if marker not in source or "WBSGEN_EnsureButtons" in source:
        return
    source = source.replace(
        marker, marker + '\n    Application.OnTime Now + TimeSerial(0, 0, 3), "ThisWorkbook.WBSGEN_EnsureButtons"', 1)
    library.replaceByName("ThisWorkbook", source + code)


def restore_vba(doc, folder):
    """シートとブックのコードを、元のテンプレートの VBA そのままに戻す。

    LibreOffice は読み込んだコードを書き換える (「Private Sub X_Click(): ...: End Sub」のような
    1 行の処理を、`Rem removed unmatched Sub/End` でコメントアウトしてしまう)。ボタンの
    クリック処理が消えるので、元のコードで置き換える。
    """
    library = doc.BasicLibraries.getByName("VBAProject")
    header = "Rem Attribute VBA_ModuleType=VBADocumentModule\nOption VBASupport 1\n"
    for name in CONTROL_MODULES + ["ThisWorkbook"]:
        path = os.path.join(folder, name + ".txt")
        if library.hasByName(name) and os.path.exists(path):
            with open(path, encoding="utf-8") as handle:
                library.replaceByName(name, header + handle.read())


def remove_controls(doc):
    """ファイルの中のボタンは外す。LibreOffice が書くボタンは Excel と食い違い、エラー
    (32809) や表示されない原因になる。ボタンは、開いたときのマクロが作る。"""
    for name in doc.Sheets.ElementNames:
        page = doc.Sheets.getByName(name).DrawPage
        dead = [page.getByIndex(i) for i in range(page.Count)
                if page.getByIndex(i).ShapeType == "com.sun.star.drawing.ControlShape"]
        for shape in dead:
            page.remove(shape)


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
    # 版番号 (Revision number)。ブックを開くときのコードが、アドインの版 (1.13) と比べて、
    # ブックの方が新しいと全ボタンを隠す。LibreOffice は保存のたびに増やすので、1 から始める
    doc.DocumentProperties.EditingCycles = 0


if __name__ == "__main__":
    main(*sys.argv[1:5])
