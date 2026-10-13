# -*- coding: utf-8 -*-
"""
sync_theme.py - excel_theme.py 의 PALETTES(단일 출처)로 Office.js 쪽 파일을 다시 만든다.

    python scripts/sync_theme.py           # theme.json · references/themes.md 재생성
    python scripts/sync_theme.py --check   # 재생성하지 않고 어긋났는지만 확인(어긋나면 종료코드 1)

색·폰트를 바꿀 때는 excel_theme.py 상단 편집 영역만 고치고 이 스크립트를 돌린다.
theme.json · references/themes.md 를 손으로 고치지 않는다(다음 동기화 때 덮어써진다).
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))          # <스킬>/scripts
ROOT = os.path.dirname(HERE)                               # <스킬>
sys.path.insert(0, HERE)
import theme_plan as tp  # noqa: E402  (excel_theme.PALETTES 를 읽는다)

THEME_JSON = os.path.join(HERE, "theme.json")
THEMES_MD = os.path.join(ROOT, "references", "themes.md")

KEYS = [("title_bg", "제목 밴드 채우기"), ("title_font", "제목 밴드 글자"),
        ("header_bg", "헤더 채우기"), ("header_text", "헤더 글자"),
        ("header2_bg", "서브표 헤더·합계 채우기"), ("header2_font", "서브표 헤더 글자"),
        ("total_fill", "합계행 채우기"), ("subheader_fill", "섹션행 채우기"),
        ("total_font", "섹션·소계·합계 글자"), ("text", "본문 글자"),
        ("border_outer", "바깥·굵은 선"), ("border_inner", "안쪽 가는 선"),
        ("input_fill", "입력 셀"), ("linked_fill", "참조 셀"), ("todo_fill", "미입수 셀")]


def theme_data():
    """패키지용 테마 데이터 - legacy 테마(navy·charcoal)는 뺀다."""
    meta = tp.list_themes()
    names = [t["name"] for t in meta["themes"] if not tp.get_theme(t["name"]).get("legacy")]
    return {
        "default": meta["default"],
        "aliases": {k: v for k, v in meta["aliases"].items() if v in names},
        "themes": {n: tp.theme_spec(n) for n in names},
        "rules": tp.rules(),
    }


def themes_md(data):
    names = list(data["themes"])
    T = data["themes"]
    L = ["# 테마별 색·서식 (scripts/sync_theme.py 가 자동 생성 - 직접 고치지 말 것)", "",
         "색은 hex(# 없이). 값을 바꾸려면 `scripts/excel_theme.py` 편집 영역(PALETTES)을 고치고 "
         "`python scripts/sync_theme.py` 를 돌린다. 스크립트를 쓸 수 있으면 스크립트를 쓴다.", "",
         "| 항목 | " + " | ".join(names) + " |", "|---|" + "---|" * len(names)]
    for k, label in KEYS:
        L.append("| %s (`%s`) | " % (label, k) + " | ".join(str(T[n]["colors"].get(k) or "-") for n in names) + " |")
    L.append("| 테두리 방식 | " + " | ".join(T[n]["border_mode"] for n in names) + " |")
    L.append("| 헤더 채우기 | " + " | ".join("예" if T[n].get("header_fill", True) else "아니오" for n in names) + " |")
    L.append("| 통화 서식 키 | " + " | ".join(T[n]["currency_format_key"] for n in names) + " |")
    L.append("| 글꼴 / 크기(제목·헤더·본문) | " + " | ".join(
        "%s %s·%s·%s" % (T[n]["font"], *[T[n]["sizes"][s] for s in ("title", "header", "body")]) for n in names) + " |")
    base = T[data["default"]] if data["default"] in T else T[names[0]]
    L += ["", "## 숫자서식 (%s 기준)" % base["name"], "", "| 키 | 서식 코드 |", "|---|---|"]
    for k, v in base["number_formats"].items():
        L.append("| `%s` | `%s` |" % (k, v.replace("|", "\\|")))
    for n in names:
        diff = {k: v for k, v in T[n]["number_formats"].items() if v != base["number_formats"].get(k)}
        if diff:
            L += ["", "%s 은 아래 서식만 다르다:" % n, ""]
            L += ["- `%s`: `%s`" % (k, v) for k, v in diff.items()]
    L += ["", "## 시트 탭 색 (역할 → hex, 빈칸 = 색 없음)", "",
          "| 역할 | " + " | ".join(names) + " |", "|---|" + "---|" * len(names)]
    for role in ("guide", "output", "calc", "input", "pbc", "raw"):
        L.append("| `%s` | " % role + " | ".join(T[n]["tab_colors"].get(role) or "" for n in names) + " |")
    r = data["rules"]
    L += ["", "규칙값: 열너비 최소 %s · 최대 %s자, 캡션 최대 %s자." % (r["min_col_width"], r["max_col_width"], r["caption_max"]), ""]
    return "\n".join(L)


def main():
    data = theme_data()
    new_json = json.dumps(data, ensure_ascii=False, indent=1) + "\n"
    new_md = themes_md(data)
    if "--check" in sys.argv:
        bad = []
        for path, new in ((THEME_JSON, new_json), (THEMES_MD, new_md)):
            try:
                with open(path, encoding="utf-8") as f:
                    if f.read() != new:
                        bad.append(path)
            except OSError:
                bad.append(path)
        if bad:
            print("어긋남 - python scripts/sync_theme.py 를 돌려 주세요:\n  " + "\n  ".join(bad))
            sys.exit(1)
        print("OK - theme.json · themes.md 가 PALETTES 와 같다 (테마 %d종)" % len(data["themes"]))
        return
    os.makedirs(os.path.dirname(THEMES_MD), exist_ok=True)
    with open(THEME_JSON, "w", encoding="utf-8", newline="\n") as f:
        f.write(new_json)
    with open(THEMES_MD, "w", encoding="utf-8", newline="\n") as f:
        f.write(new_md)
    print("동기화 완료 - 테마: " + ", ".join(data["themes"]))


if __name__ == "__main__":
    main()
