"""古い形式のスケジュール管理ブックから、入力した値を新しい形式のブックへ移す。

使い方:
    python -m tools.upgrade_workbook data/workshop_schedule.xlsx data/workshop_schedule_new.xlsx

新しいブックを tools/make_schedule_workbook.py で作り、古いブックの「入力する列」の値
(WS・タスクマスタ・メンバー・日程入力・個別タスク)を見出し名で対応させて書き写す。
数式の列(タスク一覧・自動の列)は新しいブックの数式のまま。

- 日程入力は行の位置で対応しているので、同じ行に書き写す。
- 旧形式の「作業日数(上書き)」は、開始日がある行では 終了日 = 開始日 + 日数 − 1 に変換する
  (開始日が無い行は変換できないので、書き写せなかった項目として表示する)。

作成後は Excel で一度開いて保存すること(数式の計算結果を保存するため)。
"""

import argparse
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
INPUT_SHEETS = ("WS", "タスクマスタ", "メンバー", "個別タスク")


def _headers(ws):
    """見出し名(改行の前)→ 列番号。"""

    out = {}
    for cell in ws[1]:
        if cell.value:
            out.setdefault(str(cell.value).split("\n")[0].strip(), cell.column)
    return out


def _is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def _input_columns(new_ws):
    """新しいブックで「入力する列」= 2行目が数式でない列。"""

    return {name: col for name, col in _headers(new_ws).items() if not _is_formula(new_ws.cell(2, col).value)}


def upgrade(old_path, new_path):
    subprocess.run(
        [sys.executable, "-m", "tools.make_schedule_workbook", str(new_path)], cwd=ROOT, check=True,
        capture_output=True,
    )
    old = load_workbook(old_path, data_only=True)
    new = load_workbook(new_path)
    notes = []

    # WS・マスタ・メンバー・個別タスク: 行ごとに見出し名で書き写す(新しいブックの既定マスタは消して置き換える)
    for name in INPUT_SHEETS:
        if name not in old.sheetnames:
            continue
        o, n = old[name], new[name]
        o_cols, n_cols = _headers(o), _input_columns(n)
        common = [h for h in n_cols if h in o_cols]
        for r in range(2, n.max_row + 1):
            for h in common:
                n.cell(r, n_cols[h]).value = None
        w = 2
        for r in range(2, o.max_row + 1):
            values = {h: o.cell(r, o_cols[h]).value for h in o_cols}
            if all(v in (None, "") for h, v in values.items() if h in n_cols):
                continue
            for h in common:
                n.cell(w, n_cols[h]).value = values[h]
            # 個別タスク: 旧形式は終了日が無い → 新形式でも作業日数のまま
            w += 1
        missing = [h for h in o_cols if h not in n_cols and not _is_formula(o.cell(2, o_cols[h]).value)
                   and "(自動)" not in str(o.cell(1, o_cols[h]).value)]
        if missing:
            notes.append(f"[{name}] 新しい形式に無い列は書き写していません: {', '.join(missing)}")

    # 日程入力: 同じ行に書き写す
    o, n = old["日程入力"], new["日程入力"]
    o_cols, n_cols = _headers(o), _input_columns(n)
    days_col = o_cols.get("作業日数")
    for r in range(2, o.max_row + 1):
        for h, col in n_cols.items():
            if h in o_cols:
                v = o.cell(r, o_cols[h]).value
                if v not in (None, ""):
                    n.cell(r, col).value = v
        if days_col and "終了日" in n_cols and "終了日" not in o_cols:
            days = o.cell(r, days_col).value
            start = o.cell(r, o_cols["開始日"]).value if "開始日" in o_cols else None
            if days not in (None, ""):
                if isinstance(start, datetime):
                    n.cell(r, n_cols["終了日"]).value = start + timedelta(days=int(days) - 1)
                else:
                    label = f"{o.cell(r, 1).value} / {o.cell(r, 2).value}"
                    notes.append(f"[日程入力] {r}行目({label}): 開始日が無いため作業日数 {days} を終了日に変換できません")

    new.calculation.fullCalcOnLoad = True
    new.save(new_path)
    return notes


def main():
    parser = argparse.ArgumentParser(description="古い形式のブックから入力値を新しい形式へ移す")
    parser.add_argument("old")
    parser.add_argument("new")
    args = parser.parse_args()
    notes = upgrade(args.old, args.new)
    print(f"作成しました: {args.new}(Excelで一度開いて保存してください)")
    for n in notes:
        print("  - " + n)


if __name__ == "__main__":
    main()
