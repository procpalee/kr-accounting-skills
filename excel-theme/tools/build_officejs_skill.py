# -*- coding: utf-8 -*-
"""
build_officejs_skill.py — 클로드 엑셀(Claude for Excel)용 커스텀 스킬 패키지 빌드.

원본(단일 소스)은 이 스킬 폴더다:
  PALETTES(excel_theme.py) → theme.json      plan.py(공유 알고리즘)
  officejs/officejs_plan.py · make_code.py · SKILL.template.md · references/rules.md
빌드 결과:
  dist/excel-theme-officejs/          (SKILL.md · scripts/ · references/)
  dist/excel-theme-officejs.zip       ← claude.ai > 사용자 지정 > 스킬 > 업로드

    python tools/build_officejs_skill.py            # 빌드 + 검증
    python tools/build_officejs_skill.py --no-test  # 검증(패리티·Node 목) 생략

테마·규칙을 바꾸면 이 스크립트를 다시 돌리고 ZIP 을 다시 업로드한다.
"""
import json
import os
import re
import shutil
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))          # excel-theme/tools
SKILL = os.path.dirname(HERE)                              # excel-theme
SRC = os.path.join(SKILL, "officejs")
NAME = "excel-theme-officejs"
DIST = os.path.join(SKILL, "dist")
OUT = os.path.join(DIST, NAME)
ZIP = os.path.join(DIST, NAME + ".zip")

sys.path.insert(0, SKILL)
import theme_plan as tp  # noqa: E402

LEGACY = ("navy", "charcoal")


def theme_data():
    """패키지용 테마 데이터 — legacy 테마(navy·charcoal)와 그쪽 별칭은 뺀다(사용자 확정 2026-09-14)."""
    meta = tp.list_themes()
    names = [t["name"] for t in meta["themes"] if not tp.get_theme(t["name"]).get("legacy")]
    return {
        "default": meta["default"],
        "aliases": {k: v for k, v in meta["aliases"].items() if v in names},
        "themes": {n: tp.theme_spec(n) for n in names},
        "rules": tp.rules(),
    }


def themes_md(data):
    L = ["# 테마별 색·서식 (theme.json 에서 자동 생성 — 직접 고치지 말 것)", "",
         "색은 hex(# 없이). 스크립트를 쓸 수 있으면 스크립트를 쓴다.", ""]
    keys = [("title_bg", "제목 밴드 채우기"), ("title_font", "제목 밴드 글자"),
            ("header_bg", "헤더 채우기"), ("header_text", "헤더 글자"),
            ("header2_bg", "서브표 헤더·합계 채우기"), ("header2_font", "서브표 헤더 글자"),
            ("total_fill", "합계행 채우기"), ("subheader_fill", "섹션행 채우기"),
            ("total_font", "섹션·소계·합계 글자"), ("text", "본문 글자"),
            ("border_outer", "바깥·굵은 선"), ("border_inner", "안쪽 가는 선"),
            ("input_fill", "입력 셀"), ("linked_fill", "참조 셀"), ("todo_fill", "미입수 셀")]
    names = [n for n in data["themes"] if n not in LEGACY]
    L.append("| 항목 | " + " | ".join(names) + " |")
    L.append("|---|" + "---|" * len(names))
    for k, label in keys:
        L.append("| %s (`%s`) | " % (label, k) + " | ".join(
            str(data["themes"][n]["colors"].get(k) or "-") for n in names) + " |")
    L.append("| 테두리 방식 | " + " | ".join(data["themes"][n]["border_mode"] for n in names) + " |")
    L.append("| 헤더 채우기 | " + " | ".join("예" if data["themes"][n].get("header_fill", True) else "아니오"
                                           for n in names) + " |")
    L.append("| 통화 서식 키 | " + " | ".join(data["themes"][n]["currency_format_key"] for n in names) + " |")
    L.append("| 글꼴 / 크기(제목·헤더·본문) | " + " | ".join(
        "%s %s·%s·%s" % (data["themes"][n]["font"], *[data["themes"][n]["sizes"][s] for s in ("title", "header", "body")])
        for n in names) + " |")
    L += ["", "## 숫자서식", "", "| 키 | 서식 코드 |", "|---|---|"]
    for k, v in data["themes"]["default"]["number_formats"].items():
        L.append("| `%s` | `%s` |" % (k, v.replace("|", "\\|")))
    L += ["", "dcf-valuation 은 음수에 [Red] 를 쓰지 않는다(괄호만):", ""]
    for k, v in data["themes"]["dcf-valuation"]["number_formats"].items():
        if v != data["themes"]["default"]["number_formats"].get(k):
            L.append("- `%s`: `%s`" % (k, v))
    L += ["", "## 시트 탭 색 (역할 → hex, 빈칸 = 색 없음)", "",
          "| 역할 | " + " | ".join(names) + " |", "|---|" + "---|" * len(names)]
    for role in ("guide", "output", "calc", "input", "pbc", "raw"):
        L.append("| `%s` | " % role + " | ".join(str(data["themes"][n]["tab_colors"].get(role) or "")
                                               for n in names) + " |")
    L += ["", "규칙값: 열너비 최소 %(min_col_width)s · 최대 %(max_col_width)s자, 캡션 최대 %(caption_max)s자." % data["rules"], ""]
    return "\n".join(L)


def validate_skill_md(path):
    text = open(path, encoding="utf-8").read()
    m = re.match(r"^---\n(.*?)\n---\n", text, re.S)
    if not m:
        raise SystemExit("SKILL.md 프론트매터 없음")
    fm = dict(line.split(":", 1) for line in m.group(1).splitlines() if ":" in line)
    name, desc = fm.get("name", "").strip(), fm.get("description", "").strip()
    if name != NAME:
        raise SystemExit("name(%s) ≠ 폴더명(%s)" % (name, NAME))
    if not re.fullmatch(r"[a-z0-9-]{1,64}", name):
        raise SystemExit("name 형식 오류: %s" % name)
    if not desc or len(desc) > 1024:
        raise SystemExit("description 길이 오류: %d자" % len(desc))
    return name, len(desc)


def smoke(data):
    mk = os.path.join(OUT, "scripts", "make_code.py")
    env = dict(os.environ, PYTHONIOENCODING="utf-8")
    for theme in data["themes"]:
        for mode, arg in (
            ("table", {"theme": theme, "n_cols": 3, "n_rows": 4, "currency_cols": [1], "total_rows": [4]}),
            ("parts", {"theme": theme, "parts": [{"type": "caption", "cell": "B4", "text": "표1. 잔액"},
                                                 {"type": "tab", "role": "calc"}]}),
        ):
            p = subprocess.run([sys.executable, mk, mode, json.dumps(arg, ensure_ascii=False)],
                               capture_output=True, text=True, encoding="utf-8", env=env)
            if p.returncode != 0 or "Excel.run" not in p.stdout:
                raise SystemExit("make_code 실패 (%s %s): %s" % (theme, mode, p.stderr or p.stdout))


def main():
    run_tests = "--no-test" not in sys.argv
    data = theme_data()
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(os.path.join(OUT, "scripts"))
    os.makedirs(os.path.join(OUT, "references"))

    shutil.copy(os.path.join(SRC, "SKILL.template.md"), os.path.join(OUT, "SKILL.md"))
    shutil.copy(os.path.join(SRC, "references", "rules.md"), os.path.join(OUT, "references", "rules.md"))
    with open(os.path.join(OUT, "references", "themes.md"), "w", encoding="utf-8", newline="\n") as f:
        f.write(themes_md(data))
    shutil.copy(os.path.join(SKILL, "plan.py"), os.path.join(OUT, "scripts", "plan.py"))
    for fn in ("officejs_plan.py", "make_code.py"):
        shutil.copy(os.path.join(SRC, fn), os.path.join(OUT, "scripts", fn))
    with open(os.path.join(OUT, "scripts", "theme.json"), "w", encoding="utf-8", newline="\n") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    name, dlen = validate_skill_md(os.path.join(OUT, "SKILL.md"))
    smoke(data)

    if run_tests:
        env = dict(os.environ, PYTHONIOENCODING="utf-8")
        for t in ("check_parity.py", "check_officejs.py"):
            p = subprocess.run([sys.executable, os.path.join(SRC, "tests", t)],
                               capture_output=True, text=True, encoding="utf-8", env=env)
            last = (p.stdout.strip().splitlines() or ["(no output)"])[-1]
            print("  test %-18s %s" % (t, last))
            if p.returncode != 0:
                print(p.stdout[-3000:], p.stderr[-2000:])
                raise SystemExit("테스트 실패 — ZIP 을 만들지 않았다.")

    if os.path.exists(ZIP):
        os.remove(ZIP)
    with zipfile.ZipFile(ZIP, "w", zipfile.ZIP_DEFLATED) as z:
        for root, dirs, files in os.walk(OUT):
            dirs[:] = [d for d in dirs if d != "__pycache__"]
            for fn in sorted(files):
                full = os.path.join(root, fn)
                z.write(full, os.path.relpath(full, DIST).replace(os.sep, "/"))
    with zipfile.ZipFile(ZIP) as z:
        names = z.namelist()
    print("OK  %s  (%d files, %.1f KB, description %d자, themes %d)" % (
        ZIP, len(names), os.path.getsize(ZIP) / 1024, dlen, len(data["themes"])))
    for n in names:
        print("    " + n)


if __name__ == "__main__":
    main()
