"""Excelを唯一の正(SSOT)にするスケジュール管理ブックを作る。

使い方:
    python -m tools.make_schedule_workbook data/workshop_schedule.xlsx            # 空のブック(タスクマスタのみ)
    python -m tools.make_schedule_workbook data/sample_schedule.xlsx --sample   # サンプルデータ入り

※ openpyxl で作ったファイルには数式の計算結果が入っていない。Excel(ブラウザ版含む)で
  一度開けば自動で計算される(fullCalcOnLoad)。Webアプリは計算結果を読むので、
  作成後は一度 Excel で開いて保存してから使うこと。

VBA/Office Scriptsを使わず、数式だけで
  - WSシートに種別「ワークショップ」の行を足すと、タスク一覧にマスタのタスクが並ぶ
  - 日付決定有無=Yesのタスクの日付を入れると、後続タスクの開始日・終了日が計算される
ようにする(ブラウザ版Excel・個人用OneDriveで動く範囲の関数のみ使用)。
"""

import argparse
from datetime import datetime

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

_parser = argparse.ArgumentParser(description="スケジュール管理ブック(Excel)を作る")
_parser.add_argument("output", nargs="?", default="data/workshop_schedule.xlsx")
_parser.add_argument("--sample", action="store_true", help="サンプルのWS・メンバー・日程入力を入れる")
_args = _parser.parse_args()
OUT = _args.output
SAMPLE = _args.sample

FONT = "Arial"
N_WS = 100  # WSシートの入力行数
SLOTS = 15  # 1WSあたりの日程入力の行数(=タスクマスタの最大行数)
N_MASTER = SLOTS
MAX_WS_BLOCKS = 40  # 日程入力に行を用意するワークショップ数
N_INPUT = SLOTS * MAX_WS_BLOCKS
N_OTHER = 200
N_MEMBER = 50
N_TASK = SLOTS * MAX_WS_BLOCKS  # タスク一覧に数式を入れる行数
N_ASSIGNEES = 10  # 1タスクの担当者の最大人数
STATUSES = ("未着手", "着手中", "完了")

HDR_INPUT = PatternFill("solid", start_color="305496")  # 入力するシート/列の見出し
HDR_AUTO = PatternFill("solid", start_color="548235")  # 自動計算の見出し
HDR_HELP = PatternFill("solid", start_color="808080")  # 計算用
INPUT_FILL = PatternFill("solid", start_color="FFF9E5")
AUTO_FILL = PatternFill("solid", start_color="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DATE = "yyyy/mm/dd"

wb = Workbook()


def base_font(bold=False, color="000000", size=10, italic=False):
    return Font(name=FONT, bold=bold, color=color, size=size, italic=italic)


def header(ws, cols, fills, notes=None):
    for i, (name, width) in enumerate(cols, start=1):
        c = ws.cell(row=1, column=i, value=name)
        c.font = base_font(bold=True, color="FFFFFF")
        c.fill = fills[i - 1] if isinstance(fills, list) else fills
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        ws.column_dimensions[get_column_letter(i)].width = width
        if notes and notes.get(name):
            c.comment = Comment(notes[name], "template")
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"


def body_style(ws, first_col, last_col, n_rows, fill, date_cols=(), font_color="000000"):
    for r in range(2, n_rows + 2):
        for col in range(first_col, last_col + 1):
            c = ws.cell(row=r, column=col)
            c.font = base_font(color=font_color)
            c.fill = fill
            c.border = BORDER
            if get_column_letter(col) in date_cols:
                c.number_format = DATE


def add_list(ws, rng, formula):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.showErrorMessage = True
    dv.errorTitle = "入力エラー"
    dv.error = "一覧から選んでください"
    ws.add_data_validation(dv)
    dv.add(rng)


def add_date_rule(ws, rng):
    dv = DataValidation(type="date", operator="greaterThan", formula1="1", allow_blank=True)
    dv.showErrorMessage = True
    dv.errorTitle = "入力エラー"
    dv.error = "日付を入力してください(例: 2026/10/01)"
    ws.add_data_validation(dv)
    dv.add(rng)


def add_int_rule(ws, rng):
    dv = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", allow_blank=True)
    dv.showErrorMessage = True
    dv.errorTitle = "入力エラー"
    dv.error = "1以上の整数を入力してください"
    ws.add_data_validation(dv)
    dv.add(rng)


def d(s):
    return datetime.strptime(s, "%Y/%m/%d")


# ---------------------------------------------------------------------------
# 説明
# ---------------------------------------------------------------------------
ws = wb.active
ws.title = "説明"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 100
ws["B2"] = "ワークショップ スケジュール管理ブック(このファイルが唯一の正)"
ws["B2"].font = base_font(bold=True, size=14)
lines = [
    ("このファイルの役割", ""),
    ("", "スケジュールの正式なデータはこのファイルだけです。Webアプリはこのファイルを読み取って表示するだけで、書き換えません。"),
    ("", "マクロは使っていません(ブラウザ版Excelでそのまま動きます)。タスクの追加と日付の計算は数式で自動的に行われます。"),
    ("", ""),
    ("シートの見分け方", ""),
    ("青い見出し", "入力するシート・列です(WS / 日程入力 / タスクマスタ / メンバー / 個別タスク)。薄い黄色の範囲に記入します。"),
    ("緑の見出し", "自動計算です(タスク一覧・キー・チェック列など)。直接入力しないでください。"),
    ("灰色の見出し", "計算用の列です(折りたたんで非表示にしています)。変更しないでください。"),
    ("", ""),
    ("新しいWSを追加する", ""),
    ("1", "「WS」シートの空いている行に WS名・種別(ワークショップ)・並び順・開始日 を入力します。"),
    ("2", "「タスク一覧」に、タスクマスタの「WS自動生成=対象」のタスクが自動で並びます(ボタン操作は不要)。"),
    ("", "並び順1のタスクは WS の開始日から始まり、日付決定有無=Yes のタスク(リハーサル等)以降は「未定」になります。"),
    ("", "WS の開始日が空欄だと、そのWSのタスクはすべて「未定」になります。"),
    ("", ""),
    ("日付を決める・担当者を入れる", ""),
    ("1", "「日程入力」シートにも、WSを追加した時点でそのWSのタスクの行が自動で用意されます(WS名・タスク名は自動表示)。"),
    ("2", f"該当する行の 開始日・終了日・ステータス・担当者1〜{N_ASSIGNEES}・メモ(薄い黄色の列)に入力します。"),
    ("3", "Yes のタスク(リハーサル・WS実施)に開始日を入れると、その後のタスクの日付が自動で入ります。"),
    ("4", "No のタスクに開始日を入れると「手動」になり、その日付で固定されます(空欄なら自動計算)。"),
    ("5", "終了日を入れると、作業日数ではなくその日付で終わります(開始日だけ・終了日だけ・両方、どれでも可)。"),
    ("ステータス", "未着手 / 着手中 / 完了 から選びます(空欄は未着手)。Webアプリではバーの色で表示されます。"),
    ("絞り込み", "タスク一覧・日程入力・個別タスクの見出しのフィルター(▼)で、WS名やステータスで絞り込めます。"),
    ("", f"日程入力は 1WS あたり {SLOTS} 行の枠で、WSシートのワークショップの順・タスクマスタの行の順に並びます。"),
    ("", ""),
    ("日付の計算ルール", ""),
    ("並び順1", "WS の開始日から開始。"),
    ("Yes の次", "日付決定有無=Yes のタスクの次の並び順は、その Yes タスクの終了日の翌日から開始。"),
    ("それ以外", "直前の並び順のタスクのうち最も遅い終了日の翌日から開始(直列)。同じ並び順のタスクは同じ期間(並列)。"),
    ("未定", "起点になるタスクに1件でも未定があれば、その後の自動のタスクも未定。"),
    ("警告", "Yes のタスクの開始日が前のタスクの終了日以前だと「警告」列に表示されます。"),
    ("終了日", "日程入力に終了日があればその日付、無ければ 開始日 + 標準作業日数 − 1(暦日)。"),
    ("", ""),
    ("その他の作業", ""),
    ("個別タスク", "社内MTGなど、マスタから自動で作らない作業は「個別タスク」シートに直接入力します。"),
    ("", ""),
    ("注意", ""),
    ("行の追加は必ず下に", "WS・タスクマスタは必ず一番下の空き行に追加してください。途中の行の削除・挿入・並べ替えをすると、"),
    ("", "日程入力の行とそこに入力済みの日付・担当者がずれます(入力は行の位置で対応しているため)。"),
    ("", "不要になったWSはステータスで管理し、使わなくなったマスタは「WS自動生成=対象外」にしてください(行は消さない)。"),
    ("", "ずれた場合、日程入力の「チェック」列に警告が出ます。"),
    ("タスクマスタ", "マスタの値を変更すると、すべてのWSのタスク一覧に反映されます(過去のWSも含む)。"),
    ("上限", f"タスクマスタは {SLOTS} 行まで、ワークショップは {MAX_WS_BLOCKS} 件まで。"),
]
r = 4
for label, text in lines:
    b, c = ws.cell(row=r, column=2, value=label), ws.cell(row=r, column=3, value=text)
    heading = text == "" and label != ""
    b.font = base_font(bold=True, size=12 if heading else 10, color="305496" if heading else "000000")
    c.font = base_font()
    c.alignment = Alignment(wrap_text=True, vertical="top")
    b.alignment = Alignment(vertical="top")
    r += 1
ws["B6"].fill = HDR_INPUT
ws["B6"].font = base_font(bold=True, color="FFFFFF")
ws["B7"].fill = HDR_AUTO
ws["B7"].font = base_font(bold=True, color="FFFFFF")
ws["B8"].fill = HDR_HELP
ws["B8"].font = base_font(bold=True, color="FFFFFF")

# ---------------------------------------------------------------------------
# WS
# ---------------------------------------------------------------------------
ws_ws = wb.create_sheet("WS")
header(
    ws_ws,
    [("WS名", 22), ("種別", 14), ("並び順", 9), ("開始日", 12), ("ステータス", 12), ("WS番号\n(自動)", 10)],
    [HDR_INPUT] * 5 + [HDR_AUTO],
    {
        "WS名": "重複不可。日程入力・個別タスクはこの名前で紐づけます",
        "種別": "ワークショップ = タスクマスタから自動でタスクを作る / その他 = 作らない",
        "開始日": "並び順1のタスクの開始日になります",
    },
)
body_style(ws_ws, 1, 5, N_WS, INPUT_FILL, date_cols=("D",))
body_style(ws_ws, 6, 6, N_WS, AUTO_FILL)
for r in range(2, N_WS + 2):
    ws_ws[f"F{r}"] = (
        f'=IF(AND($A{r}<>"",$B{r}="ワークショップ"),'
        f'COUNTIFS($A$2:$A{r},"<>",$B$2:$B{r},"ワークショップ"),"")'
    )
ws_rows = [
    ["WS 秋季研修", "ワークショップ", 1, d("2026/10/01"), "準備中"],
    ["Snow Dorm", "ワークショップ", 2, d("2026/11/02"), "企画中"],
    ["その他", "その他", 99, None, None],
]
for i, row in enumerate(ws_rows if SAMPLE else [], start=2):
    for j, v in enumerate(row, start=1):
        ws_ws.cell(row=i, column=j, value=v)
add_list(ws_ws, f"B2:B{N_WS + 1}", '"ワークショップ,その他"')
add_date_rule(ws_ws, f"D2:D{N_WS + 1}")
add_int_rule(ws_ws, f"C2:C{N_WS + 1}")

# ---------------------------------------------------------------------------
# タスクマスタ
# ---------------------------------------------------------------------------
ws_m = wb.create_sheet("タスクマスタ")
header(
    ws_m,
    [("タスク名", 22), ("標準作業日数", 12), ("標準並び順", 11), ("日付決定有無", 12), ("WS自動生成", 12), ("生成順\n(自動)", 10), ("日程入力の\n行位置(自動)", 12)],
    [HDR_INPUT] * 5 + [HDR_AUTO] * 2,
    {
        "標準並び順": "同じ値のタスクは並列(同じ期間)になります",
        "日付決定有無": "Yes = 人が日付を決めるタスク。決まるまで開始日は未定で、以降のタスクも未定になります",
        "WS自動生成": "空欄 = 対象。対象外にすると新しいWSに作られません",
    },
)
body_style(ws_m, 1, 5, N_MASTER, INPUT_FILL)
body_style(ws_m, 6, 7, N_MASTER, AUTO_FILL)
last = N_MASTER + 1
for r in range(2, last + 1):
    # 生成順 = 標準並び順の小さい順(同じ並び順なら上の行から)の通し番号
    ws_m[f"F{r}"] = (
        f'=IF(AND($A{r}<>"",$E{r}<>"対象外"),'
        f'COUNTIFS($A$2:$A${last},"<>",$E$2:$E${last},"<>対象外",$C$2:$C${last},"<"&$C{r})'
        f'+COUNTIFS($A$2:$A{r},"<>",$E$2:$E{r},"<>対象外",$C$2:$C{r},$C{r}),"")'
    )
    # 日程入力の各WSの枠の中での行位置 = マスタの行の順(下に追加しても既存の行はずれない)
    ws_m[f"G{r}"] = f'=IF(AND($A{r}<>"",$E{r}<>"対象外"),ROW()-1,"")'
masters = [
    ("コンテンツ開発", 28, 1, "No"),
    ("場所決め", 28, 1, "No"),
    ("リハーサル物品購入", 28, 1, "No"),
    ("リハーサル&広報撮影", 1, 2, "Yes"),
    ("募集フォーム作成", 14, 3, "No"),
    ("広報動画作成", 14, 3, "No"),
    ("LP作成", 14, 3, "No"),
    ("広報稼働", 28, 4, "No"),
    ("物品購入", 28, 4, "No"),
    ("ワークショップ実施", 1, 5, "Yes"),
    ("アンケート分析", 7, 6, "No"),
]
for i, (name, days, order, yes) in enumerate(masters, start=2):
    for j, v in enumerate([name, days, order, yes, "対象"], start=1):
        ws_m.cell(row=i, column=j, value=v)
add_list(ws_m, f"D2:D{last}", '"Yes,No"')
add_list(ws_m, f"E2:E{last}", '"対象,対象外"')
add_int_rule(ws_m, f"B2:C{last}")

# ---------------------------------------------------------------------------
# メンバー
# ---------------------------------------------------------------------------
ws_mem = wb.create_sheet("メンバー")
header(
    ws_mem,
    [("姓", 10), ("名", 10), ("表示名", 16), ("メールアドレス", 26), ("状態", 8)],
    HDR_INPUT,
    {"表示名": "重複不可。担当者のプルダウンに出る名前です"},
)
body_style(ws_mem, 1, 5, N_MEMBER, INPUT_FILL)
for i, row in enumerate(
    [] if not SAMPLE else [
        ["佐藤", "太郎", "佐藤 太郎", "sato@example.com", "有効"],
        ["鈴木", "次郎", "鈴木 次郎", "suzuki@example.com", "有効"],
        ["田中", "三郎", "田中 三郎", "tanaka@example.com", "有効"],
        ["村上", "四郎", "村上 四郎", "murakami@example.com", "有効"],
    ],
    start=2,
):
    for j, v in enumerate(row, start=1):
        ws_mem.cell(row=i, column=j, value=v)
add_list(ws_mem, f"E2:E{N_MEMBER + 1}", '"有効,無効"')

# ---------------------------------------------------------------------------
# 日程入力(左の4列は自動、右の列に入力する)
# ---------------------------------------------------------------------------
# 行 k(=ROW()-1)は「k番目のワークショップ」×「タスクマスタの k 番目の行」に固定で対応する。
#   枠番号 = INT((k-1)/SLOTS)+1 → WSシートの WS番号(ワークショップの通し番号)
#   枠内位置 = MOD(k-1,SLOTS)+1 → タスクマスタの行位置
# WS・マスタを下に追加しても既存の行の位置は変わらないので、入力済みの値がずれない。
L = get_column_letter
IN_START, IN_END, IN_STATUS = "E", "F", "G"
IN_AS1, IN_ASN = L(8), L(7 + N_ASSIGNEES)  # 担当者1〜N
IN_MEMO, IN_KEY, IN_CHECK = L(8 + N_ASSIGNEES), L(9 + N_ASSIGNEES), L(10 + N_ASSIGNEES)
IN_H = [L(12 + N_ASSIGNEES + i) for i in range(4)]  # 枠番号・枠内位置・WS行・マスタ行
ws_in = wb.create_sheet("日程入力")
header(
    ws_in,
    [("WS名\n(自動)", 20), ("タスク名\n(自動)", 22), ("並び順\n(自動)", 8), ("日付決定有無\n(自動)", 11),
     ("開始日", 12), ("終了日", 12), ("ステータス", 10)]
    + [(f"担当者{i}", 12) for i in range(1, N_ASSIGNEES + 1)]
    + [("メモ", 24), ("キー\n(自動)", 30), ("チェック\n(自動)", 34), ("", 2),
       ("枠番号", 8), ("枠内位置", 8), ("WS行", 7), ("マスタ行", 8)],
    [HDR_AUTO] * 4 + [HDR_INPUT] * (4 + N_ASSIGNEES) + [HDR_AUTO] * 2 + [PatternFill()] + [HDR_HELP] * 4,
    {
        "開始日": "Yesのタスク: 決定した日付 / Noのタスク: 入れると手動で固定、空欄なら自動計算",
        "終了日": "入れるとその日付で終了(空欄なら 開始日+標準作業日数−1)",
        "ステータス": "未着手 / 着手中 / 完了(空欄は未着手)",
    },
)
ws_in.freeze_panes = "C2"
lin = N_INPUT + 1
n_in_cols = 10 + N_ASSIGNEES
body_style(ws_in, 1, 4, N_INPUT, AUTO_FILL)
body_style(ws_in, 5, 8 + N_ASSIGNEES, N_INPUT, INPUT_FILL, date_cols=(IN_START, IN_END))
body_style(ws_in, 9 + N_ASSIGNEES, n_in_cols, N_INPUT, AUTO_FILL)
body_style(ws_in, 12 + N_ASSIGNEES, 15 + N_ASSIGNEES, N_INPUT, AUTO_FILL, font_color="595959")
TMq = "'タスクマスタ'"
HB, HS, HW, HM = IN_H
for r in range(2, lin + 1):
    ws_in[f"{HB}{r}"] = f"=INT((ROW()-2)/{SLOTS})+1"
    ws_in[f"{HS}{r}"] = f"=MOD(ROW()-2,{SLOTS})+1"
    ws_in[f"{HW}{r}"] = f'=IFERROR(MATCH(${HB}{r},WS!$F$1:$F${N_WS + 1},0),"")'
    ws_in[f"{HM}{r}"] = f'=IF(${HW}{r}="","",IFERROR(MATCH(${HS}{r},{TMq}!$G$1:$G${N_MASTER + 1},0),""))'
    ws_in[f"A{r}"] = f'=IF(${HM}{r}="","",INDEX(WS!$A$1:$A${N_WS + 1},${HW}{r}))'
    ws_in[f"B{r}"] = f'=IF(${HM}{r}="","",INDEX({TMq}!$A$1:$A${N_MASTER + 1},${HM}{r}))'
    ws_in[f"C{r}"] = f'=IF(${HM}{r}="","",INDEX({TMq}!$C$1:$C${N_MASTER + 1},${HM}{r}))'
    ws_in[f"D{r}"] = f'=IF(${HM}{r}="","",IF(INDEX({TMq}!$D$1:$D${N_MASTER + 1},${HM}{r})="Yes","Yes","No"))'
    ws_in[f"{IN_KEY}{r}"] = f'=IF($A{r}="","",$A{r}&"|"&$B{r})'
    ws_in[f"{IN_CHECK}{r}"] = (
        f'=IF($A{r}="",IF(COUNTA(${IN_START}{r}:${IN_MEMO}{r})>0,"対応するタスクがありません(入力が無視されます)",""),'
        f'IF(AND(ISNUMBER(${IN_START}{r}),ISNUMBER(${IN_END}{r})),IF(${IN_END}{r}<${IN_START}{r},"終了日が開始日より前です",""),"")'
        f'&IF(AND($D{r}="Yes",${IN_START}{r}=""),"日付未決定",IF(AND($D{r}="No",${IN_START}{r}<>""),"手動で固定","")))'
    )
    for col in ("C", "D", IN_STATUS):
        ws_in[f"{col}{r}"].alignment = Alignment(horizontal="center")

# サンプルの入力(WS 秋季研修 = 1〜15行目の枠、Snow Dorm = 16〜30行目の枠)
slot_of = {name: i for i, (name, *_rest) in enumerate(masters, start=1)}


def input_row(ws_no, task):
    return 1 + (ws_no - 1) * SLOTS + slot_of[task]


# (WS番号, タスク名, 開始日, 終了日, ステータス, [担当者], メモ)
inputs = [
    (1, "コンテンツ開発", None, None, "完了", ["田中 三郎"], "担当者だけ入力(日付は自動)"),
    (1, "場所決め", d("2026/09/28"), d("2026/10/11"), "完了", ["村上 四郎"], "開始日・終了日を手動で指定"),
    (1, "リハーサル物品購入", None, None, "着手中", ["鈴木 次郎"], ""),
    (1, "リハーサル&広報撮影", d("2026/11/10"), None, None, ["佐藤 太郎", "鈴木 次郎"], "日付決定"),
    (1, "ワークショップ実施", d("2026/12/20"), None, None, ["佐藤 太郎", "鈴木 次郎", "田中 三郎", "村上 四郎"],
     "日付決定(広報稼働と重なるため警告が出る例)"),
    (2, "コンテンツ開発", None, None, "着手中", ["鈴木 次郎"], ""),
]
for ws_no, task, start, end, status, names, memo in inputs if SAMPLE else []:
    r = input_row(ws_no, task)
    ws_in[f"{IN_START}{r}"], ws_in[f"{IN_END}{r}"], ws_in[f"{IN_STATUS}{r}"] = start, end, status
    for i, n in enumerate(names):
        ws_in.cell(row=r, column=8 + i, value=n)
    ws_in[f"{IN_MEMO}{r}"] = memo or None

add_list(ws_in, f"{IN_STATUS}2:{IN_STATUS}{lin}", '"' + ",".join(STATUSES) + '"')
add_list(ws_in, f"{IN_AS1}2:{IN_ASN}{lin}", f"=メンバー!$C$2:$C${N_MEMBER + 1}")
add_date_rule(ws_in, f"{IN_START}2:{IN_END}{lin}")
# タスクの無い行はグレー、WSの枠ごとに縞模様、チェックの警告は赤
ws_in.conditional_formatting.add(
    f"A2:{IN_CHECK}{lin}",
    FormulaRule(formula=['$A2=""'], fill=PatternFill("solid", start_color="D9D9D9"),
                font=Font(name=FONT, color="A6A6A6")),
)
ws_in.conditional_formatting.add(
    f"A2:D{lin}",
    FormulaRule(formula=[f'AND($A2<>"",ISODD(${HB}2))'], fill=PatternFill("solid", start_color="EAF1FB")),
)
ws_in.conditional_formatting.add(
    f"{IN_CHECK}2:{IN_CHECK}{lin}",
    FormulaRule(formula=[f'OR(LEFT(${IN_CHECK}2,4)="対応する",LEFT(${IN_CHECK}2,4)="終了日が")'],
                fill=PatternFill("solid", start_color="F8CBAD"), font=Font(name=FONT, color="C00000", bold=True)),
)
ws_in.conditional_formatting.add(
    f"{IN_CHECK}2:{IN_CHECK}{lin}",
    FormulaRule(formula=[f'${IN_CHECK}2="日付未決定"'], font=Font(name=FONT, color="C55A11", bold=True)),
)
STATUS_FILLS = {"未着手": "FFF2CC", "着手中": "DDEBF7", "完了": "D9D9D9"}
for st, color in STATUS_FILLS.items():
    ws_in.conditional_formatting.add(
        f"{IN_STATUS}2:{IN_STATUS}{lin}",
        FormulaRule(formula=[f'${IN_STATUS}2="{st}"'], fill=PatternFill("solid", start_color=color)),
    )
ws_in.column_dimensions.group(L(11 + N_ASSIGNEES), IN_H[-1], hidden=True, outline_level=1)
ws_in.auto_filter.ref = f"A1:{IN_CHECK}{lin}"

# ---------------------------------------------------------------------------
# 個別タスク
# ---------------------------------------------------------------------------
ws_o = wb.create_sheet("個別タスク")
header(
    ws_o,
    [("WS名", 20), ("タスク名", 22), ("開始日", 12), ("終了日", 12), ("作業日数", 10), ("ステータス", 10)]
    + [(f"担当者{i}", 12) for i in range(1, N_ASSIGNEES + 1)]
    + [("メモ", 24)],
    HDR_INPUT,
    {
        "WS名": "マスタから自動で作らない作業(社内MTGなど)をここに直接書きます",
        "終了日": "空欄なら 開始日+作業日数−1",
        "ステータス": "未着手 / 着手中 / 完了(空欄は未着手)",
    },
)
lo = N_OTHER + 1
o_last = 7 + N_ASSIGNEES
body_style(ws_o, 1, o_last, N_OTHER, INPUT_FILL, date_cols=("C", "D"))
for i, row in enumerate(
    [] if not SAMPLE else [
        ["その他", "社内MTG", d("2026/10/06"), None, 1, "完了", "佐藤 太郎", "鈴木 次郎", "田中 三郎"],
        ["その他", "資料作成", d("2026/10/13"), d("2026/10/16"), None, "未着手", "田中 三郎"],
    ],
    start=2,
):
    for j, v in enumerate(row, start=1):
        ws_o.cell(row=i, column=j, value=v)
add_list(ws_o, f"A2:A{lo}", f"=WS!$A$2:$A${N_WS + 1}")
add_list(ws_o, f"F2:F{lo}", '"' + ",".join(STATUSES) + '"')
add_list(ws_o, f"G2:{L(6 + N_ASSIGNEES)}{lo}", f"=メンバー!$C$2:$C${N_MEMBER + 1}")
add_date_rule(ws_o, f"C2:D{lo}")
add_int_rule(ws_o, f"E2:E{lo}")
for st, color in STATUS_FILLS.items():
    ws_o.conditional_formatting.add(
        f"F2:F{lo}", FormulaRule(formula=[f'$F2="{st}"'], fill=PatternFill("solid", start_color=color)),
    )
ws_o.auto_filter.ref = f"A1:{L(o_last)}{lo}"

# ---------------------------------------------------------------------------
# タスク一覧(すべて数式)
# ---------------------------------------------------------------------------
ws_t = wb.create_sheet("タスク一覧", 1)
cols = [
    ("No", 6), ("WS名", 20), ("タスク名", 22), ("並び順", 8), ("作業日数", 9), ("日付決定有無", 11),
    ("開始日の決め方", 12), ("開始日", 12), ("終了日", 12), ("担当者", 30), ("ステータス", 10),
    ("警告", 34), ("メモ", 24),
    ("", 2),
    ("WS連番", 8), ("マスタ連番", 9), ("WS行", 7), ("マスタ行", 8), ("入力行", 7), ("入力開始日", 11),
    ("入力終了日", 11), ("WS開始日", 11), ("前の並び順", 9), ("起点日", 11), ("キー", 30), ("標準日数", 8),
    ("", 2), ("設定", 16), ("値", 8),
]
fills = [HDR_AUTO] * 13 + [PatternFill()] + [HDR_HELP] * 12 + [PatternFill(), HDR_HELP, HDR_HELP]
header(
    ws_t,
    cols,
    fills,
    {
        "開始日の決め方": "決定待ち / 決定 = Yesのタスク。手動 = 日程入力で開始日を指定。自動 = ルールで計算",
        "警告": "Yesのタスクの開始日が前のタスクの終了日以前のとき、終了日が開始日より前のときに表示",
    },
)
ws_t.freeze_panes = "D2"
ws_t["N1"].fill = PatternFill()
ws_t["AA1"].fill = PatternFill()
ws_t["AB2"] = "対象マスタ数"
ws_t["AC2"] = f"=COUNT(タスクマスタ!$F$2:$F${N_MASTER + 1})"
ws_t["AB3"] = "ワークショップ数"
ws_t["AC3"] = f"=MAX(WS!$F$2:$F${N_WS + 1})"
for c in ("AB2", "AC2", "AB3", "AC3"):
    ws_t[c].font = base_font(color="595959")
    ws_t[c].fill = AUTO_FILL
    ws_t[c].border = BORDER

body_style(ws_t, 1, 13, N_TASK, PatternFill(), date_cols=("H", "I"))
body_style(ws_t, 15, 26, N_TASK, AUTO_FILL, date_cols=("T", "U", "V", "X"), font_color="595959")

M, W = "$AC$2", "$AC$3"
IN = "'日程入力'"
TM = "'タスクマスタ'"
NI = N_INPUT + 1
for r in range(2, N_TASK + 2):
    p = r - 1  # 直前の行(範囲は見出し行1から直前の行まで=自分の行を含めない)
    rng = lambda col: f"${col}$1:${col}{p}"  # noqa: E731
    f = {}
    # --- 計算用 ---
    f["O"] = f'=IF({M}=0,"",IF(INT((ROW()-2)/{M})+1>{W},"",INT((ROW()-2)/{M})+1))'
    f["P"] = f'=IF($O{r}="","",MOD(ROW()-2,{M})+1)'
    f["Q"] = f'=IF($O{r}="","",MATCH($O{r},WS!$F$1:$F${N_WS + 1},0))'
    f["R"] = f'=IF($P{r}="","",MATCH($P{r},{TM}!$F$1:$F${N_MASTER + 1},0))'
    f["Y"] = f'=IF($B{r}="","",$B{r}&"|"&$C{r})'
    f["S"] = f'=IF($Y{r}="","",IFERROR(MATCH($Y{r},{IN}!${IN_KEY}$1:${IN_KEY}${NI},0),""))'
    for col, src in (("T", IN_START), ("U", IN_END)):
        f[col] = (
            f'=IF($S{r}="","",IF(INDEX({IN}!${src}$1:${src}${NI},$S{r})="","",'
            f"INDEX({IN}!${src}$1:${src}${NI},$S{r})))"
        )
    f["V"] = (
        f'=IF($Q{r}="","",IF(INDEX(WS!$D$1:$D${N_WS + 1},$Q{r})="","未定",'
        f"INDEX(WS!$D$1:$D${N_WS + 1},$Q{r})))"
    )
    f["Z"] = f'=IF($R{r}="","",INDEX({TM}!$B$1:$B${N_MASTER + 1},$R{r}))'
    f["W"] = (
        f'=IF($B{r}="","",IF(COUNTIFS({rng("B")},$B{r},{rng("D")},"<"&$D{r})=0,"なし",'
        f'_xlfn.MAXIFS({rng("D")},{rng("B")},$B{r},{rng("D")},"<"&$D{r})))'
    )
    same_prev = f'{rng("B")},$B{r},{rng("D")},$W{r}'
    f["X"] = (
        f'=IF($B{r}="","",IF($W{r}="なし",$V{r},'
        f'IF(COUNTIFS({same_prev},{rng("F")},"Yes")>0,'
        f'IF(COUNTIFS({same_prev},{rng("F")},"Yes",{rng("H")},"未定")>0,"未定",'
        f'_xlfn.MAXIFS({rng("I")},{same_prev},{rng("F")},"Yes")+1),'
        f'IF(COUNTIFS({same_prev},{rng("H")},"未定")>0,"未定",'
        f'_xlfn.MAXIFS({rng("I")},{same_prev})+1))))'
    )
    # --- 表示列 ---
    f["A"] = f'=IF($B{r}="","",COUNTIF($B$2:$B{r},"?*"))'
    f["B"] = f'=IF($Q{r}="","",INDEX(WS!$A$1:$A${N_WS + 1},$Q{r}))'
    f["C"] = f'=IF($R{r}="","",INDEX({TM}!$A$1:$A${N_MASTER + 1},$R{r}))'
    f["D"] = f'=IF($R{r}="","",INDEX({TM}!$C$1:$C${N_MASTER + 1},$R{r}))'
    f["F"] = f'=IF($R{r}="","",IF(INDEX({TM}!$D$1:$D${N_MASTER + 1},$R{r})="Yes","Yes","No"))'
    f["G"] = (
        f'=IF($B{r}="","",IF($F{r}="Yes",IF($T{r}="","決定待ち","決定"),'
        f'IF($T{r}<>"","手動",IF($U{r}<>"","自動(終了日指定)","自動"))))'
    )
    f["H"] = f'=IF($B{r}="","",IF($F{r}="Yes",IF($T{r}="","未定",$T{r}),IF($T{r}<>"",$T{r},$X{r})))'
    f["I"] = f'=IF($B{r}="","",IF($U{r}<>"",$U{r},IF(ISNUMBER($H{r}),$H{r}+$Z{r}-1,"未定")))'
    f["E"] = f'=IF($B{r}="","",IF(AND(ISNUMBER($H{r}),ISNUMBER($I{r})),$I{r}-$H{r}+1,$Z{r}))'
    f["J"] = f'=IF($S{r}="","",_xlfn.TEXTJOIN("、",TRUE,INDEX({IN}!${IN_AS1}$1:${IN_ASN}${NI},$S{r},0)))'
    f["K"] = (
        f'=IF($B{r}="","",IF($S{r}="","未着手",IF(INDEX({IN}!${IN_STATUS}$1:${IN_STATUS}${NI},$S{r})="","未着手",'
        f"INDEX({IN}!${IN_STATUS}$1:${IN_STATUS}${NI},$S{r}))))"
    )
    f["L"] = (
        f'=IF($B{r}="","",IF(AND($W{r}="なし",$V{r}="未定"),"WSの開始日が未入力です",'
        f'IF(AND(ISNUMBER($H{r}),ISNUMBER($I{r}),$I{r}<$H{r}),"終了日が開始日より前です",'
        f'IF(AND($F{r}="Yes",ISNUMBER($H{r}),ISNUMBER($X{r})),'
        f'IF($H{r}<$X{r},"前のタスクの終了日("&TEXT($X{r}-1,"yyyy/mm/dd")&")以前です",""),""))))'
    )
    f["M"] = (
        f'=IF($S{r}="","",IF(INDEX({IN}!${IN_MEMO}$1:${IN_MEMO}${NI},$S{r})="","",'
        f"INDEX({IN}!${IN_MEMO}$1:${IN_MEMO}${NI},$S{r})))"
    )
    for col, formula in f.items():
        ws_t[f"{col}{r}"] = formula
    for col in ("A", "D", "E", "F", "G", "H", "I", "K"):
        ws_t[f"{col}{r}"].alignment = Alignment(horizontal="center")

last_t = N_TASK + 1
# WSごとに縞模様、未定・決定待ち・警告・ステータスを色分け
ws_t.conditional_formatting.add(
    f"A2:M{last_t}",
    FormulaRule(formula=['AND($B2<>"",ISODD($O2))'], fill=PatternFill("solid", start_color="EAF1FB")),
)
ws_t.conditional_formatting.add(
    f"H2:I{last_t}",
    FormulaRule(formula=['H2="未定"'], font=Font(name=FONT, color="808080", italic=True)),
)
ws_t.conditional_formatting.add(
    f"G2:G{last_t}",
    FormulaRule(formula=['G2="決定待ち"'], font=Font(name=FONT, color="C55A11", bold=True)),
)
ws_t.conditional_formatting.add(
    f"L2:L{last_t}",
    FormulaRule(formula=['L2<>""'], fill=PatternFill("solid", start_color="F8CBAD"),
                font=Font(name=FONT, color="C00000", bold=True)),
)
for st, color in STATUS_FILLS.items():
    ws_t.conditional_formatting.add(
        f"K2:K{last_t}", FormulaRule(formula=[f'$K2="{st}"'], fill=PatternFill("solid", start_color=color)),
    )
# 計算用の列は折りたたんで隠す
ws_t.column_dimensions.group("N", "AC", hidden=True, outline_level=1)
ws_t.auto_filter.ref = f"A1:M{last_t}"

# シートの並び: 説明 / タスク一覧 / WS / 日程入力 / タスクマスタ / メンバー / 個別タスク
order = ["説明", "タスク一覧", "WS", "日程入力", "タスクマスタ", "メンバー", "個別タスク"]
wb._sheets = [wb[n] for n in order]
wb["タスク一覧"].sheet_properties.tabColor = "548235"
for n in ("WS", "日程入力", "タスクマスタ", "メンバー", "個別タスク"):
    wb[n].sheet_properties.tabColor = "305496"
wb.active = 1
wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("saved", OUT)
