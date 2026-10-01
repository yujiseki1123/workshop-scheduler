"""JSON API・画面のテスト(データは Excel のスケジュール管理ブックから読む)。フィクスチャは conftest.py。"""

from tests.conftest import TASK_ROWS, d, fill, other_row


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------


def test_schedule_api_reads_excel(book_path, client_for):
    client = client_for(book_path)
    payload = client.get("/api/schedule").get_json()
    assert [p["project_name"] for p in payload["projects"]] == ["WS 秋", "Snow Dorm", "その他"]
    assert len(payload["tasks"]) == len(TASK_ROWS) + 2
    assert payload["source"]["path"].endswith("book.xlsx")
    assert payload["warnings"]


def test_schedule_api_reports_missing_file(tmp_path, client_for):
    client = client_for(tmp_path / "none.xlsx")
    res = client.get("/api/schedule")
    assert res.status_code == 503
    assert "見つかりません" in res.get_json()["error"]


def test_members_projects_masters_api(book_path, client_for):
    client = client_for(book_path)
    assert [m["display_name"] for m in client.get("/api/members").get_json()][:3] == [
        "佐藤 太郎", "鈴木 次郎", "田中 三郎",
    ]
    assert client.get("/api/projects").get_json()[0]["start_date"] == "2026-10-01"
    assert len(client.get("/api/task-masters").get_json()) == 11


def test_reload_when_file_changes(book_path, client_for):
    client = client_for(book_path)
    assert len(client.get("/api/schedule").get_json()["tasks"]) == len(TASK_ROWS) + 2
    fill(book_path, other_rows=[
        other_row("その他", "社内MTG", d("2026/10/06"), days=1),
        other_row("その他", "資料作成", d("2026/10/13"), days=3),
        other_row("その他", "合宿準備", d("2026/10/20"), days=2),
    ])
    assert len(client.get("/api/schedule").get_json()["tasks"]) == len(TASK_ROWS) + 3


def test_schedule_page_renders(book_path, client_for):
    res = client_for(book_path).get("/")
    assert res.status_code == 200 and b"schedule-root" in res.data


def test_schedule_api_returns_status(book_path, client_for):
    tasks = client_for(book_path).get("/api/schedule").get_json()["tasks"]
    assert {t["status"] for t in tasks} == {"未着手", "着手中", "完了"}
