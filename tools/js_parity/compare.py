"""Python 版と JavaScript 版の Excel 読み取り結果が同じか確かめる。

使い方(初回だけ npm install が必要。Node.js が要る):
    cd tools/js_parity && npm install && cd ../..
    python -m tools.js_parity.compare data/workshop_schedule.xlsx [他のブック...]

excel_source.py か docs/js/excel_source.js のどちらかを直したら、両方に同じ修正を入れてこれで確認する。
"""

import json
import subprocess
import sys
import warnings
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))

from app.services.excel_source import parse_workbook  # noqa: E402


def python_result(path):
    try:
        book = parse_workbook(str(path))
    except Exception as e:  # エラーの種類と文言も一致させる
        return {"error": f"{type(e).__name__}: {e}"}
    data = {
        "projects": [{k: p[k] for k in ("project_id", "project_name", "project_category")} for p in book.projects],
        "tasks": book.tasks,
        "warnings": book.warnings,
        "source": None,
    }
    return json.loads(json.dumps({"data": data, "members": book.members, "task_masters": book.task_masters},
                                 ensure_ascii=False, default=str))


def js_result(path):
    out = subprocess.run(["node", str(HERE / "dump.js"), str(path)], capture_output=True, text=True, check=True)
    return json.loads(out.stdout)


def main(paths):
    warnings.filterwarnings("ignore")
    ok = True
    for path in paths:
        py, js = python_result(path), js_result(path)
        if py == js:
            n = len(py["data"]["tasks"]) if "data" in py else py["error"]
            print(f"一致  {path}({n})")
            continue
        ok = False
        print(f"不一致 {path}")
        if "error" in py or "error" in js:
            print("  Python:", py.get("error"), "\n  JS    :", js.get("error"))
            continue
        for key in ("projects", "warnings"):
            if py["data"][key] != js["data"][key]:
                print(f"  {key}\n    Python: {py['data'][key]}\n    JS    : {js['data'][key]}")
        for i, (a, b) in enumerate(zip(py["data"]["tasks"], js["data"]["tasks"])):
            if a != b:
                print(f"  tasks[{i}]", {k: (a.get(k), b.get(k)) for k in sorted(set(a) | set(b)) if a.get(k) != b.get(k)})
        if len(py["data"]["tasks"]) != len(js["data"]["tasks"]):
            print("  タスク数", len(py["data"]["tasks"]), len(js["data"]["tasks"]))
        for key in ("members", "task_masters"):
            if py[key] != js[key]:
                print(f"  {key} が違います")
    return 0 if ok else 1


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1:]))
