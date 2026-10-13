# -*- coding: utf-8 -*-
"""
export_theme.py — PALETTES(단일 소스) → 웹 MCP가 쓸 theme.json 생성 + plan.py 복사.

웹 MCP(excel-theme-mcp)는 로컬 PALETTES를 직접 import 하지 않고(배포 분리),
이 스크립트가 만든 theme.json(데이터) + plan.py(알고리즘)만 쓴다.
→ 테마를 바꾸면 이 스크립트를 한 번 돌리고 commit/push 하면 웹이 따라온다.
  (tools/pre-commit 훅을 걸면 commit 시 자동 실행)

    python tools/export_theme.py
"""
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))          # excel-theme/tools
SKILL = os.path.dirname(HERE)                              # excel-theme
SKILLS_ROOT = os.path.dirname(SKILL)                       # skills/
MCP = os.path.join(SKILLS_ROOT, "excel-theme-mcp")

sys.path.insert(0, SKILL)
import theme_plan as tp  # noqa: E402


def main():
    os.makedirs(MCP, exist_ok=True)
    meta = tp.list_themes()
    data = {
        "default": meta["default"],
        "aliases": meta["aliases"],
        "themes": {t["name"]: tp.theme_spec(t["name"]) for t in meta["themes"]},
        "rules": tp.rules(),
    }
    with open(os.path.join(MCP, "theme.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    shutil.copy(os.path.join(SKILL, "plan.py"), os.path.join(MCP, "plan.py"))
    print("OK  theme.json (%d themes) + plan.py -> %s" % (len(data["themes"]), MCP))


if __name__ == "__main__":
    main()
