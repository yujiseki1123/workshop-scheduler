"""JSON API。

フロントエンドのJS(static/js/schedule.js)がfetchで呼び出す。
データは Excel のスケジュール管理ブックから読む(読み取り専用)。
"""

from flask import Blueprint, jsonify

from app.services.excel_source import ExcelSourceError
from app.services.schedule_service import get_book, get_schedule_data

api_bp = Blueprint("api", __name__)


@api_bp.errorhandler(ExcelSourceError)
def excel_error(e):
    # ファイルが無い・形式が違う・計算結果が無い などは画面にそのまま表示する
    return jsonify({"error": str(e)}), 503


@api_bp.get("/schedule")
def schedule():
    """スケジュール画面が必要とする projects + tasks をまとめて返す。"""

    return jsonify(get_schedule_data())


@api_bp.get("/members")
def members():
    return jsonify(get_book().members)


@api_bp.get("/projects")
def projects():
    return jsonify(
        [
            {k: (v.isoformat() if hasattr(v, "isoformat") else v) for k, v in p.items()}
            for p in get_book().projects
        ]
    )


@api_bp.get("/task-masters")
def task_masters():
    return jsonify(get_book().task_masters)
