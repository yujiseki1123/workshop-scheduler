"""HTML画面を返すルート。"""

from datetime import date

from flask import Blueprint, current_app, render_template

views_bp = Blueprint("views", __name__)


@views_bp.get("/")
def schedule_page():
    """スケジュール画面。実データはJSがAPI(/api/schedule)から取得する。

    テンプレートには「今日の日付」と「何週間分表示するか」だけを渡し、
    タイムラインの週割り自体はフロントJS側で組み立てる。
    """

    return render_template(
        "schedule.html",
        today=date.today().isoformat(),
        weeks_ahead=current_app.config["SCHEDULE_WEEKS_AHEAD"],
        weeks_behind=current_app.config["SCHEDULE_WEEKS_BEHIND"],
    )
