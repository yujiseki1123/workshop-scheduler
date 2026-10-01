"""スケジュール画面・APIのための業務ロジック層。

データの取得元は Excel のスケジュール管理ブック(唯一の正)。
読み取りは app/services/excel_source.py に任せ、ここでは画面が必要とする形に整える。
ルート(routes/api.py)は HTTP の入出力に専念させる。
"""

from flask import current_app

from app.services.excel_source import load_book


def excel_path():
    return current_app.config["SCHEDULE_EXCEL_PATH"]


def get_book():
    return load_book(excel_path())


def get_schedule_data() -> dict:
    """スケジュール画面の描画に必要な全データ。

    projects(行)と tasks(バー)を分けて返す。開始日が未定のタスクは
    start_date / end_date が null。warnings はブックの読み取り時に見つかった問題、
    タスクごとの警告(日付の前後関係など)は各タスクの warning に入る。
    """

    book = get_book()
    return {
        "projects": [
            {
                "project_id": p["project_id"],
                "project_name": p["project_name"],
                "project_category": p["project_category"],
            }
            for p in book.projects
        ],
        "tasks": book.tasks,
        "warnings": book.warnings,
        "source": {"path": book.path, "modified_at": book.modified_at},
    }
