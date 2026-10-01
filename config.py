"""アプリケーション設定。

PoCでは環境変数をほぼ使わずシンプルな固定値で済ませるが、将来的に
本番/開発/テストで設定を切り替えやすいよう、クラスベースの設定にしている。
"""

import os

# プロジェクトのルートディレクトリ(このファイルがある場所)
BASE_DIR = os.path.abspath(os.path.dirname(__file__))

# スケジュールの唯一の正となる Excel ブック。環境変数 SCHEDULE_EXCEL_PATH で差し替え可能
# (例: OneDrive の同期フォルダ内のファイルを直接指定する)。
EXCEL_PATH = os.environ.get(
    "SCHEDULE_EXCEL_PATH", os.path.join(BASE_DIR, "data", "workshop_schedule.xlsx")
)


class Config:
    """共通のデフォルト設定。"""

    SCHEDULE_EXCEL_PATH = EXCEL_PATH
    # スケジュール画面が初期表示時に何週間先まで用意するか(横スクロール分)
    SCHEDULE_WEEKS_AHEAD = 10
    # スケジュール画面が現在日から何週間前まで表示するか
    SCHEDULE_WEEKS_BEHIND = 1


class TestConfig(Config):
    """テスト用設定。"""

    TESTING = True
