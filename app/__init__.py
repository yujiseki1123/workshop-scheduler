"""アプリケーションファクトリ。

create_app() を使うことで、本番用アプリとテスト用アプリ(設定違い)を
同じコードから作り分けられるようにしている(Flaskの定石パターン)。
データは Excel のスケジュール管理ブックから読むため、DBは使わない。
"""

from flask import Flask

from config import Config


def create_app(config_class=Config):
    app = Flask(__name__)
    app.config.from_object(config_class)

    # Blueprintの登録。View(HTML)とAPI(JSON)でルートを分離している。
    from app.routes.views import views_bp
    from app.routes.api import api_bp

    app.register_blueprint(views_bp)
    app.register_blueprint(api_bp, url_prefix="/api")

    return app
