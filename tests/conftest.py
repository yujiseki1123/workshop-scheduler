"""pytest共通フィクスチャ。

ブックは tools/make_schedule_workbook.py で実際に作り、タスク一覧の自動計算の列(開始日など)は
「Excelが保存した計算結果」に見立てた値で上書きする(テスト環境では数式を計算できないため)。
"""

import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import load_workbook

from app import create_app
from app.services.excel_source import clear_cache
from config import TestConfig

# ---------------------------------------------------------------------------

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture(scope="session")
def empty_book(tmp_path_factory):
    path = tmp_path_factory.mktemp("tpl") / "empty.xlsx"
    subprocess.run(
        [sys.executable, "-m", "tools.make_schedule_workbook", str(path)], cwd=ROOT, check=True,
        capture_output=True,
    )
    return path


def d(s):  # noqa: D103
    return datetime.strptime(s, "%Y/%m/%d")


# タスク一覧の値(WS名・タスク名など)と計算結果(開始日など)。列は見出し名で対応させる(TASK_COLUMNS)
TASK_COLUMNS = ["No", "WS名", "タスク名", "並び順", "作業日数", "日付決定有無", "開始日の決め方", "開始日", "終了日",
                "担当者", "ステータス", "警告", "メモ"]
TASK_ROWS = [
    # No, WS名, タスク名, 並び順, 作業日数, 日付決定有無, 開始日の決め方, 開始日, 終了日, 担当者, ステータス, 警告, メモ
    [1, "WS 秋", "コンテンツ開発", 1, 28, "No", "自動", d("2026/10/01"), d("2026/10/28"), "田中 三郎", "完了", None, "担当者だけ"],
    [2, "WS 秋", "リハーサル&広報撮影", 2, 1, "Yes", "決定", d("2026/11/10"), d("2026/11/10"), "佐藤 太郎、鈴木 次郎",
     "着手中", None, None],
    [3, "WS 秋", "LP作成", 3, 14, "No", "自動", d("2026/11/11"), d("2026/11/24"), None, "未着手", None, None],
    [4, "WS 秋", "ワークショップ実施", 5, 1, "Yes", "決定", d("2026/12/20"), d("2026/12/20"), None, "未着手",
     "前のタスクの終了日(2026/12/22)以前です", None],
    [5, "Snow Dorm", "コンテンツ開発", 1, 28, "No", "自動", d("2026/11/02"), d("2026/11/29"), "外部 花子", "未着手", None, None],
    [6, "Snow Dorm", "リハーサル&広報撮影", 2, 1, "Yes", "決定待ち", "未定", "未定", None, "未着手", None, None],
    [7, "Snow Dorm", "LP作成", 3, 14, "No", "自動", "未定", "未定", None, None, None, None],
]

TEN = [f"担当{i}" for i in range(1, 11)]


def other_row(ws, name, start=None, end=None, days=None, status=None, names=(), memo=None):
    """個別タスクの1行(WS名, タスク名, 開始日, 終了日, 作業日数, ステータス, 担当者1〜10, メモ)。"""

    return [ws, name, start, end, days, status] + list(names) + [None] * (10 - len(names)) + [memo]


def headers(ws):
    return {str(c.value).split("\n")[0]: c.column for c in ws[1] if c.value}


def fill(path, ws_rows=(), task_rows=(), other_rows=(), members=(), task_columns=TASK_COLUMNS):
    wb = load_workbook(path)
    for sheet, rows in (("WS", ws_rows), ("個別タスク", other_rows), ("メンバー", members)):
        ws = wb[sheet]
        for i, row in enumerate(rows, start=2):
            for j, v in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=v)
    t = wb["タスク一覧"]
    cols = headers(t)
    for i, row in enumerate(task_rows, start=2):
        for name, v in zip(task_columns, row):
            if name in cols:  # 「No」は今の形式には無い
                t.cell(row=i, column=cols[name], value=v)
    wb.save(path)


@pytest.fixture
def book_path(empty_book, tmp_path):
    path = tmp_path / "book.xlsx"
    shutil.copy(empty_book, path)
    fill(
        path,
        ws_rows=[
            ["Snow Dorm", "ワークショップ", 2, d("2026/11/02"), "企画中"],
            ["WS 秋", "ワークショップ", 1, d("2026/10/01"), "準備中"],
            ["その他", "その他", 99, None, None],
        ],
        task_rows=TASK_ROWS,
        other_rows=[
            other_row("その他", "社内MTG", d("2026/10/06"), None, 1, "完了", ["佐藤 太郎", "鈴木 次郎"], "定例"),
            other_row("存在しないWS", "打合せ", d("2026/10/07"), None, 1),
            other_row("その他", "合宿準備", d("2026/10/13"), d("2026/10/16"), None, "着手中", TEN),
        ],
        members=[
            ["佐藤", "太郎", "佐藤 太郎", "sato@example.com", "有効"],
            ["鈴木", "次郎", "鈴木 次郎", "suzuki@example.com", "有効"],
            ["田中", "三郎", "田中 三郎", "tanaka@example.com", "有効"],
        ],
    )
    return path


@pytest.fixture
def client_for(tmp_path):
    def make(path):
        clear_cache()

        class Cfg(TestConfig):
            SCHEDULE_EXCEL_PATH = str(path)

        return create_app(Cfg).test_client()

    return make


