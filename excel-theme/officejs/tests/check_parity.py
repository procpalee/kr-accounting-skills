# -*- coding: utf-8 -*-
"""
check_parity.py — 스킬 패키지(excel-theme-officejs)의 Office.js ops 가 로컬 excel-theme 스킬 결과와 같은지 셀 단위로 대조한다.

A = 스킬 원본: excel_theme.write_table / write_title / write_audit_header / write_section_bar /
    write_caption / set_tab 으로 만든 openpyxl 시트
B = 같은 값만 넣은 빈 시트에 officejs_plan 의 ops 를 Excel 규칙대로 적용한 결과(해석기)

대조 항목: 값 · 채우기 · 글꼴(이름·크기·굵게·색) · 숫자서식 · 정렬(가로·세로·줄바꿈) ·
          보이는 테두리선(가로·세로, 굵기+색) · A열 너비 · 탭 색.
대조 제외: 자동 열너비(Excel 렌더링 의존) · 워크북 Normal 스타일.

빌드(tools/build_officejs_skill.py)가 dist 를 만든 뒤 자동 실행한다.
    python officejs/tests/check_parity.py
"""
import json
import os
import re
import sys

SKILL = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # excel-theme
ROOT = os.path.join(SKILL, "dist", "excel-theme-officejs", "scripts")   # 빌드된 스킬 패키지
sys.path.insert(0, ROOT)
sys.path.insert(0, SKILL)

from openpyxl import Workbook                               # noqa: E402
from excel_theme import (write_table, write_title, write_audit_header,  # noqa: E402
                         write_section_bar, write_caption, set_tab, _horiz_for, sheet_label)
from officejs_plan import officejs_plan, officejs_sheet_parts  # noqa: E402

with open(os.path.join(ROOT, "theme.json"), encoding="utf-8") as f:
    DATA = json.load(f)
THEMES, RULES = DATA["themes"], DATA["rules"]
STRENGTH = {"hairline": 1, "thin": 2, "medium": 3, "double": 4, "thick": 5}
DASH = re.compile("[—–―]")


def col_num(s):
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


def parse(addr):
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", addr)
    c1, r1 = col_num(m.group(1)), int(m.group(2))
    return r1, c1, int(m.group(4) or r1), col_num(m.group(3) or m.group(1))


def rgb(c):
    if c is None:
        return None
    v = getattr(c, "rgb", c)
    return v[-6:].upper() if isinstance(v, str) else None


# ── B: Office.js ops 해석기 ─────────────────────────────────────
class Sheet:
    def __init__(self, title):
        self.title = title
        self.cells = {}
        self.edges = {}
        self.colA = None
        self.tab = None

    def c(self, r, k):
        return self.cells.setdefault((r, k), {"value": None, "fill": None, "bold": False, "color": "000000",
                                             "size": 11, "name": "Calibri", "nf": "General",
                                             "h": None, "v": None, "wrap": False})

    def each(self, addr):
        r1, c1, r2, c2 = parse(addr)
        for r in range(r1, r2 + 1):
            for k in range(c1, c2 + 1):
                yield r, k


def edge_keys(addr, edge):
    r1, c1, r2, c2 = parse(addr)
    if edge == "EdgeTop":
        return [("h", r1, k) for k in range(c1, c2 + 1)]
    if edge == "EdgeBottom":
        return [("h", r2 + 1, k) for k in range(c1, c2 + 1)]
    if edge == "EdgeLeft":
        return [("v", r, c1) for r in range(r1, r2 + 1)]
    if edge == "EdgeRight":
        return [("v", r, c2 + 1) for r in range(r1, r2 + 1)]
    if edge == "InsideHorizontal":
        return [("h", r, k) for r in range(r1 + 1, r2 + 1) for k in range(c1, c2 + 1)]
    if edge == "InsideVertical":
        return [("v", r, k) for r in range(r1, r2 + 1) for k in range(c1 + 1, c2 + 1)]
    raise ValueError(edge)


def apply(sheet, ops):
    label = re.sub(r"^\d+\s+", "", sheet.title).strip()
    for o in ops:
        k, a = o["op"], o.get("address")
        if k in ("workbookFont", "unfreeze", "autofitColumns", "negativeHighlight"):
            continue
        if k == "columnWidth":
            sheet.colA = o["points"]
        elif k == "fill":
            for r, c in sheet.each(a):
                sheet.c(r, c)["fill"] = o["color"][1:]
        elif k == "font":
            for r, c in sheet.each(a):
                cell = sheet.c(r, c)
                for p, key in (("name", "name"), ("size", "size"), ("bold", "bold")):
                    if p in o:
                        cell[key] = o[p]
                if "color" in o:
                    cell["color"] = o["color"][1:]
        elif k == "horizontalAlignment":
            for r, c in sheet.each(a):
                sheet.c(r, c)["h"] = o["value"]
        elif k == "verticalAlignment":
            for r, c in sheet.each(a):
                sheet.c(r, c)["v"] = o["value"].lower()
        elif k == "wrapText":
            for r, c in sheet.each(a):
                sheet.c(r, c)["wrap"] = o["value"]
        elif k == "numberFormat":
            for r, c in sheet.each(a):
                sheet.c(r, c)["nf"] = o["format"]
        elif k == "border":
            for key in edge_keys(a, o["edge"]):
                if o["style"] == "None":
                    sheet.edges.pop(key, None)
                else:
                    w = "double" if o["style"] == "Double" else o["weight"].lower()
                    sheet.edges[key] = (w, o["color"][1:])
        elif k == "normalizePunct":
            for r, c in sheet.each(a):
                cell = sheet.c(r, c)
                v = cell["value"]
                if isinstance(v, str) and not v.startswith("=") and DASH.search(v):
                    cell["value"] = DASH.sub("-", v)
        elif k == "value":
            for r, c in sheet.each(a):
                sheet.c(r, c)["value"] = label if o.get("sheetLabel") else o["text"]
        elif k == "tabColor":
            sheet.tab = o["color"][1:] if o["color"] else None
        else:
            raise ValueError("interpreter: unknown op %s" % k)


def resolve_h(cell):
    h = cell["h"]
    if h is None:
        return None
    if h == "General":
        return _horiz_for(cell["value"])
    return h.lower()


# ── A: openpyxl 에서 같은 항목 추출 ─────────────────────────────
def a_cell(ws, r, k):
    c = ws.cell(row=r, column=k)
    fill = rgb(c.fill.fgColor) if c.fill is not None and c.fill.fill_type == "solid" else None
    al = c.alignment
    return {"value": c.value, "fill": fill, "bold": bool(c.font.b), "color": rgb(c.font.color) or "000000",
            "size": float(c.font.sz or 11), "name": c.font.name, "nf": c.number_format,
            "h": al.horizontal, "v": al.vertical, "wrap": bool(al.wrap_text)}


def a_edge(ws, kind, r, k):
    if kind == "h":
        sides = [ws.cell(row=r - 1, column=k).border.bottom if r > 1 else None,
                 ws.cell(row=r, column=k).border.top]
    else:
        sides = [ws.cell(row=r, column=k - 1).border.right if k > 1 else None,
                 ws.cell(row=r, column=k).border.left]
    sides = [(s.style, rgb(s.color)) for s in sides if s is not None and s.style]
    if not sides:
        return None
    best = max(STRENGTH[s[0]] for s in sides)
    top = [s for s in sides if STRENGTH[s[0]] == best]
    colors = {s[1] for s in top}
    return (top[0][0], top[0][1] if len(colors) == 1 else "*")


def compare(ws, sheet, max_row, max_col, tag):
    diffs = []
    for r in range(1, max_row + 1):
        for k in range(1, max_col + 1):
            A = a_cell(ws, r, k)
            B = dict(sheet.cells.get((r, k)) or Sheet("x").c(r, k))
            B["h"] = resolve_h(B)
            B["size"] = float(B["size"])
            for key in ("value", "fill", "bold", "color", "size", "name", "nf", "h", "v", "wrap"):
                if A[key] != B[key]:
                    diffs.append("%s %s%d %s: skill=%r server=%r" % (tag, chr(64 + k), r, key, A[key], B[key]))
            for kind in ("h", "v"):
                ea, eb = a_edge(ws, kind, r, k), sheet.edges.get((kind, r, k))
                if ea is None and eb is None:
                    continue
                if ea is None or eb is None or ea[0] != eb[0] or (ea[1] != "*" and ea[1] != eb[1]):
                    diffs.append("%s %s%d edge-%s: skill=%r server=%r" % (tag, chr(64 + k), r, kind, ea, eb))
    wa = ws.column_dimensions["A"].width
    if sheet.colA is not None and abs(wa - 2.0) > 1e-9:
        diffs.append("%s colA width skill=%r" % (tag, wa))
    ta = rgb(ws.sheet_properties.tabColor) if ws.sheet_properties.tabColor is not None else None
    if ta != sheet.tab:
        diffs.append("%s tab: skill=%r server=%r" % (tag, ta, sheet.tab))
    return diffs


# ── 시나리오 ──────────────────────────────────────────────────
HEADERS = ["계정 — 구분", "당기", "전기", "증감률"]
ROWS = [
    ["영업활동", None, None, None],
    ["매출액", 1200000000, 1000000000, 0.2],
    ["매출원가 – 제품", -800000000, -700000000, 0.1428],
    ["매출총이익", "=C6+C7", "=D6+D7", "=IF(D8=0,\"\",C8/D8-1)"],
    ["판매비", -100000000, -90000000, 0.11],
    ["영업이익", "=C8+C9", "=D8+D9", 0.3],
]
TABLES = [
    dict(tag="기본", kw=dict(currency_cols=[1, 2], percent_cols=[3], section_rows=[1], subtotal_rows=[4],
                          total_rows=[6], input_cells=["C6:D7"], linked_cells=["C9:D9"], todo_cells=["E6"])),
    dict(tag="서브표", kw=dict(currency_cols=[1, 2], number_cols={"E": "percent"}, subtotal_rows=[2, 4],
                            total_rows=[5], secondary=True)),
    dict(tag="연속합계", kw=dict(number_cols={1: "accounting", 2: "accounting"}, total_rows=[3],
                              subtotal_rows=[4], header_row=6, start_col=2)),
    dict(tag="소계끝행", kw=dict(currency_cols=[1], subtotal_rows=[6], section_rows=[1])),
]


def table_case(theme, case):
    kw = dict(case["kw"])
    header_row = kw.pop("header_row", 4)
    start_col = kw.pop("start_col", 2)
    title = None if theme == "audit" else "재무상태표 — 요약"
    wb = Workbook(); ws = wb.active; ws.title = "10 재무상태표"
    write_table(ws, HEADERS, ROWS, theme=theme, title=title, header_row=header_row,
                start_col=start_col, **kw)

    sheet = Sheet(ws.title)
    if title is not None:
        sheet.c(2, 2)["value"] = title
    for j, h in enumerate(HEADERS):
        sheet.c(header_row, start_col + j)["value"] = h
    for i, row in enumerate(ROWS):
        for j, v in enumerate(row):
            if v is not None:
                sheet.c(header_row + 1 + i, start_col + j)["value"] = v
    number_cols = kw.pop("number_cols", None)
    plan = officejs_plan(THEMES[theme], rules=RULES, header_row=header_row, start_col=start_col,
                         n_cols=len(HEADERS), n_rows=len(ROWS), title_cell="auto",
                         number_format_cols=number_cols, **kw)
    apply(sheet, plan["ops"])
    return compare(ws, sheet, header_row + len(ROWS) + 2, start_col + len(HEADERS) + 1,
                   "%s/%s" % (theme, case["tag"]))


def parts_case(theme):
    wb = Workbook(); ws = wb.active; ws.title = "20 연령분석"
    meta = {"회사명": "A사", "조서번호": "6000A-20", "결산일": "2026-12-31",
            "조서명": "매출채권 — 연령분석", "작성자": "작성자"}
    if theme == "audit":
        write_audit_header(ws, meta, theme=theme, tab="output")
        parts = [{"type": "audit_header", "meta": meta}, {"type": "tab", "role": "결과"}]
    else:
        write_title(ws, theme=theme, last_col=7, tab="calc")
        parts = [{"type": "title_band", "last_col": 7}, {"type": "tab", "role": "계산"}]
    write_section_bar(ws, 6, "1. 총괄표 — 요약", theme=theme, last_col=8)
    write_caption(ws, "B9", "표1. 잔액 – 연령별", theme=theme)
    parts += [{"type": "section_bar", "row": 6, "text": "1. 총괄표 — 요약", "last_col": 8},
              {"type": "caption", "cell": "B9", "text": "표1. 잔액 – 연령별"}]
    res = officejs_sheet_parts(THEMES[theme], parts, rules=RULES)
    sheet = Sheet(ws.title)
    apply(sheet, res["ops"])
    diffs = compare(ws, sheet, 11, 9, "%s/시트요소" % theme)
    want_next = 11  # caption B9 → 표 헤더 11행
    if res["next_header_row"] != want_next:
        diffs.append("%s/시트요소 next_header_row=%r (기대 %d)" % (theme, res["next_header_row"], want_next))
    return diffs


def main():
    fails = 0
    for theme in THEMES:
        for case in TABLES:
            d = table_case(theme, case)
            fails += len(d)
            print("%-4s %-14s %-8s %s" % ("OK" if not d else "DIFF", theme, case["tag"],
                                           "" if not d else "(%d)" % len(d)))
            for line in d[:12]:
                print("      " + line)
        d = parts_case(theme)
        fails += len(d)
        print("%-4s %-14s %-8s %s" % ("OK" if not d else "DIFF", theme, "시트요소", "" if not d else "(%d)" % len(d)))
        for line in d[:12]:
            print("      " + line)
    print("\n%s (%d diff)" % ("PARITY OK" if not fails else "PARITY FAILED", fails))
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
