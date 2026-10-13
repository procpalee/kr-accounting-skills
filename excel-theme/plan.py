# -*- coding: utf-8 -*-
"""
plan.py — 순수 format_plan: 테마 spec(dict) → Excel 서식 연산 리스트(중립 ops).

의존성 없음(표준 라이브러리만) → 그대로 웹 MCP 서버에 복사해 쓸 수 있다.
excel-theme 스킬(theme_plan.py)과 excel-theme-mcp(웹) 가 **이 한 파일**을 공유한다.
(export_theme.py 가 이 파일을 MCP 폴더로 복사 + theme.json 생성)

★ 기준은 excel_theme.write_table / apply_theme 의 실제 결과다(2026-09-13 동기화).
   ops 를 순서대로 적용하면 write_table 로 만든 표와 셀 서식이 같아야 한다
   (excel-theme-mcp/tests/check_parity.py 가 셀 단위로 대조한다).
   - 행높이는 지정하지 않는다(사용자 확정 2026-09-09).
   - 틀고정은 두지 않는다(발견 시 해제).
   - 긴 대시(— – ―) → "-" (수식 제외, 2026-09-11).
   - 열너비: 자동맞춤 + 상한 MAX_COL_WIDTH(30).
   - 합계/소계 테두리는 excel_theme.style_total_row 의 셀 양면 규칙을 그대로 재현한 뒤
     "눈에 보이는 선"(위 셀 아래선·아래 셀 윗선 중 굵은 쪽, 같으면 나중에 칠한 쪽)으로 풀어 낸다.

sheet_parts_plan 은 표 밖 시트 요소(제목 밴드·조서 헤더·섹션 바·캡션·탭 색)를 같은 방식으로 만든다.
"""
import re

MAX_COL_WIDTH = 30          # excel_theme.MAX_COL_WIDTH (rules 로 덮어씀)
MIN_COL_WIDTH = 8           # excel_theme.autofit_columns 기본 min_width
CAPTION_MAX = 40            # theme_lint.MAX_CAPTION

_DASH_RE = re.compile("[—–―]")
_TITLE_TAIL = re.compile(r"\s*(?:[—–\-|:·]|\(|\[|/)\s.*$|\s*[—–|]\s*.*$")
_STRENGTH = {"hairline": 1, "thin": 2, "medium": 3, "double": 4, "thick": 5}
_BLACK = "000000"


# ── 주소 유틸 ─────────────────────────────────────────────────
def _letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _rng(c1, r1, c2, r2):
    return "%s%d:%s%d" % (_letter(c1), r1, _letter(c2), r2)


def _cell(c, r):
    return "%s%d" % (_letter(c), r)


def _split_cell(ref):
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError("cell reference expected (e.g. B2): %s" % ref)
    return _col_num(m.group(1)), int(m.group(2))


def _coll(start_col, key):
    return key if isinstance(key, str) else _letter(start_col + key)


# ── 텍스트 규칙 (excel_theme.plain_punct / title_only 와 동일) ──────────
def plain_punct(value):
    if not isinstance(value, str) or value.startswith("="):
        return value
    return _DASH_RE.sub("-", value)


def title_only(text):
    """제목 밴드 문자열 → 시트명 한 덩어리. 비었으면 None(= 실행 시 시트명 사용)."""
    if text is None or not str(text).strip():
        return None
    t = re.sub(r"^\d+\s+", "", str(text)).strip()
    t = _TITLE_TAIL.sub("", t).strip()
    return t or None


def _rules(rules):
    r = {"max_col_width": MAX_COL_WIDTH, "min_col_width": MIN_COL_WIDTH,
         "caption_max": CAPTION_MAX, "tab_role_aliases": {}}
    r.update(rules or {})
    return r


def _font(rng, spec, size, bold, color):
    return {"op": "font", "range": rng, "name": spec["font"], "size": size,
            "bold": bold, "color": color or _BLACK}


def _align(rng, horizontal, vertical="center", wrap=False):
    return {"op": "align", "range": rng, "horizontal": horizontal,
            "vertical": vertical, "wrap": wrap}


# ── 표 ────────────────────────────────────────────────────────
def format_plan(spec, *, header_row=4, start_col=2, n_cols=1, n_rows=1, title_cell="auto",
                number_format_cols=None, currency_cols=(), percent_cols=(),
                section_rows=(), subtotal_rows=(), total_rows=(),
                input_cells=(), linked_cells=(), todo_cells=(),
                secondary=False, autofit=True, highlight_neg=None, rules=None):
    """테마 spec + 표 레이아웃 → 결정론적 Excel 서식 연산 리스트 (write_table 과 같은 결과).

    인덱스 규약: 헤더=header_row, 데이터=그 아래, 표는 start_col(기본 B=2)부터.
    section/subtotal/total_rows = 데이터 1-based 인덱스. number_format_cols 키=열문자 또는 0-based 데이터열.
    title_cell: "auto" = B2 (audit 테마는 제목 밴드 없음) · "" = 제목 없음 · "B2" 등 직접 지정.
    secondary : 서브표 — 헤더·합계 배경을 header2 색으로.
    """
    rl = _rules(rules)
    col = spec["colors"]
    sizes = spec["sizes"]
    fmts = spec["number_formats"]
    mode = spec["border_mode"]
    oc, ic = col["border_outer"], col["border_inner"]
    header_fill = spec.get("header_fill", True)
    total_font = col.get("total_font")
    end_col = start_col + n_cols - 1
    data_start = header_row + 1
    last_row = data_start + n_rows - 1
    table = _rng(start_col, header_row, end_col, last_row)
    body = _rng(start_col, data_start, end_col, last_row)
    hdr = _rng(start_col, header_row, end_col, header_row)
    row_rng = lambda r: _rng(start_col, r, end_col, r)

    ops = [{"op": "workbook_font", "name": spec["font"], "size": sizes["body"]},
           {"op": "unfreeze"},
           {"op": "column_width", "col": "A", "width": 2.0}]

    # 제목(B2) — 색을 데이터 끝열까지 확장
    if title_cell == "auto":
        title_cell = "" if spec["name"] == "audit" else "B2"
    if title_cell:
        tcol, trow = _split_cell(title_cell)
        ops.append({"op": "normalize_punct", "range": title_cell})
        if spec.get("title_fill"):
            ops.append({"op": "fill", "range": _rng(tcol, trow, max(end_col, tcol), trow),
                        "color": col["title_bg"]})
            ops.append(_font(_rng(tcol, trow, max(end_col, tcol), trow), spec, sizes["title"], True,
                             col["title_font"]))   # 밴드 전체에 글꼴(excel_theme.apply_theme 와 동일)
        else:
            ops.append(_font(title_cell, spec, sizes["title"], True,
                             col.get("title_color", col["title_font"])))
        ops.append(_align(title_cell, "left"))

    ops.append({"op": "normalize_punct", "range": table})

    # 본문 — 숫자·수식 오른쪽 / 텍스트 왼쪽(by_type), 세로 가운데
    ops.append(_font(body, spec, sizes["body"], False, col["text"]))
    ops.append(_align(body, "by_type"))

    # 헤더
    if secondary:
        hbg, htxt = col["header2_bg"], col["header2_font"]
    else:
        hbg, htxt = col["header_bg"], col.get("header_text", col["header_font"])
    if header_fill:
        ops.append({"op": "fill", "range": hdr, "color": hbg})
    ops.append(_font(hdr, spec, sizes["header"], True, htxt))
    ops.append(_align(hdr, "center", wrap=True))

    # 테두리 — 표 기본선 + 합계/소계 선을 "보이는 선" 기준으로 계산
    ops.extend(_border_ops(spec, mode, oc, ic, header_fill, header_row, last_row,
                           table, row_rng, subtotal_rows, total_rows, data_start))

    # 숫자/통화/퍼센트 서식 (열 단위, 오른쪽 정렬)
    cf = {}
    for k in currency_cols:
        cf[_coll(start_col, k)] = spec["currency_format_key"]
    for k in percent_cols:
        cf[_coll(start_col, k)] = "percent_acct"
    for k, v in (number_format_cols or {}).items():
        cf[_coll(start_col, k)] = v
    for cl, fk in cf.items():
        ops.append({"op": "number_format", "range": "%s%d:%s%d" % (cl, data_start, cl, last_row),
                    "format": fmts.get(fk, fk), "align": "right"})

    # 음수 강조: 테마가 굵게 강조면 조건부서식, 아니면 숫자서식 [Red]만
    if highlight_neg is None:
        highlight_neg = spec.get("highlight_neg_bold", False)
    if highlight_neg:
        ops.append({"op": "conditional_negative", "range": body,
                    "color": spec["accent"], "bold": True})

    if autofit:
        ops.append({"op": "autofit_columns", "range": table,
                    "min_width": rl["min_col_width"], "max_width": rl["max_col_width"]})

    # 섹션행: subheader_fill + 굵게
    for idx in section_rows:
        rg = row_rng(data_start + idx - 1)
        ops.append({"op": "fill", "range": rg, "color": col["subheader_fill"]})
        ops.append(_font(rg, spec, sizes["body"], True, total_font))
    # 소계(배경 없음) → 합계(배경)
    for idx in subtotal_rows:
        ops.append(_font(row_rng(data_start + idx - 1), spec, sizes["body"], True, total_font))
    for idx in total_rows:
        rg = row_rng(data_start + idx - 1)
        ops.append(_font(rg, spec, sizes["body"], True, total_font))
        ops.append({"op": "fill", "range": rg,
                    "color": col["header2_bg"] if secondary else col["total_fill"]})

    # 입력(크림) / 연결 / 미입수(노랑)
    for rg in (input_cells or []):
        ops.append({"op": "fill", "range": rg, "color": col["input_fill"]})
    for rg in (linked_cells or []):
        ops.append({"op": "fill", "range": rg, "color": col["linked_fill"]})
    for rg in (todo_cells or []):
        ops.append({"op": "fill", "range": rg, "color": col.get("todo_fill", "FFFF00")})

    return {"theme": spec["name"], "ops": ops}


def _border_ops(spec, mode, oc, ic, header_fill, header_row, last_row, table, row_rng,
                subtotal_rows, total_rows, data_start):
    """excel_theme(style_header_row → apply_borders → style_total_row) 의 셀 양면 테두리를 재현해
    가로선마다 '보이는 선'을 구하고, 표 기본 op 가 그리는 선과 다른 곳만 border_edges 로 덧칠한다."""
    rows = range(header_row, last_row + 1)
    top, bot = {}, {}
    if mode in ("frame", "grid"):
        for r in rows:
            top[r] = ("medium", oc, 0) if r == header_row else ("thin", ic, 0)
            bot[r] = ("medium", oc, 0) if r == last_row else ("thin", ic, 0)
    elif mode == "statement":
        top[header_row] = ("medium", oc, 0)
        bot[header_row] = ("thin", oc, 0)
        bot[last_row] = ("medium", oc, 0)
    elif mode == "rules" and not header_fill:
        bot[header_row] = ("medium", oc, 0)

    total_style = spec.get("total_style", "default")
    rules_like = mode in ("rules", "statement")
    seq = [0]

    def style_row(r, fill):
        seq[0] += 1
        if total_style == "double":
            t, b = ("thin", oc), (("double" if fill else "thin"), oc)
        elif rules_like:
            t, b = ("thin", oc), (("medium", oc) if fill else None)
        else:
            t, b = ("medium", oc), ("thin", oc)
        top[r] = t + (seq[0],)
        prev = bot.get(r)
        keep = prev is not None and prev[0] in ("medium", "thick", "double")
        if not keep and b is not None:
            bot[r] = b + (seq[0],)

    for idx in subtotal_rows:
        style_row(data_start + idx - 1, False)
    for idx in total_rows:
        style_row(data_start + idx - 1, True)

    def visual(k):
        cands = [s for s in (bot.get(k - 1), top.get(k)) if s is not None]
        if not cands:
            return None
        s = max(cands, key=lambda x: (_STRENGTH[x[0]], x[2]))
        return (s[0], s[1])

    def op_visual(k):                       # 표 기본 op(border) 가 그리는 선
        if mode in ("frame", "grid"):
            return ("medium", oc) if k in (header_row, last_row + 1) else ("thin", ic)
        if mode == "statement":
            return ("medium", oc) if k in (header_row, last_row + 1) else None
        return None

    ops = []
    if mode == "frame":
        ops.append({"op": "border", "range": table, "style": "frame", "sides": "top_bottom",
                    "outer_color": oc, "outer_weight": "medium",
                    "inner_color": ic, "inner_weight": "thin"})
    elif mode == "grid":
        ops.append({"op": "border", "range": table, "style": "grid",
                    "outer_color": oc, "outer_weight": "medium",
                    "inner_color": ic, "inner_weight": "thin"})
    elif mode == "statement":
        ops.append({"op": "border", "range": table, "style": "statement", "sides": "top_bottom",
                    "outer_color": oc, "outer_weight": "medium"})
    for k in range(header_row, last_row + 2):
        want = visual(k)
        if want == op_visual(k):
            continue
        side = {"style": "none"} if want is None else {"style": want[0], "color": want[1]}
        if k <= last_row:
            ops.append({"op": "border_edges", "range": row_rng(k), "top": side})
        else:
            ops.append({"op": "border_edges", "range": row_rng(last_row), "bottom": side})
    return ops


# ── 시트 요소 (표 밖) ──────────────────────────────────────────
AUDIT_HEADER_FIELDS = ("회사명", "조서번호", "시트명", "결산일", "조서명", "작성자")


def sheet_parts_plan(spec, parts, rules=None):
    """시트 요소 목록 → ops. parts 항목(type 별):
      {"type":"title_band","first_col":2,"last_col":7,"text":None}   # text 생략=시트명(권장)
      {"type":"audit_header","meta":{회사명,조서번호,결산일,조서명,작성자[,시트명]},"first_col":2}
      {"type":"section_bar","row":6,"text":"1. 총괄표","first_col":2,"last_col":8}
      {"type":"caption","cell":"B12","text":"표1. …"}
      {"type":"tab","role":"calc"}   # guide/output/calc/input/pbc/raw (+한글 별칭)
    반환: {theme, ops, next_header_row, warnings}
    """
    rl = _rules(rules)
    col, sizes = spec["colors"], spec["sizes"]
    oc, ic = col["border_outer"], col["border_inner"]
    ops, warnings, next_row = [], [], None
    body = sizes["body"]

    for p in parts:
        t = p.get("type")
        if t == "title_band":
            first, last = p.get("first_col", 2), p.get("last_col", 7)
            if spec["name"] == "audit":
                warnings.append("audit 테마는 제목 밴드를 두지 않는다 — audit_header 를 쓴다.")
            band = _rng(first, 2, last, 2)
            ops.append({"op": "column_width", "col": "A", "width": 2.0})
            ops.append({"op": "fill", "range": band, "color": col["title_bg"]})
            ops.append(_font(band, spec, sizes["title"], True, col["title_font"]))
            text = title_only(plain_punct(p.get("text")))
            cell = _cell(first, 2)
            ops.append({"op": "value", "cell": cell, "text": text} if text
                       else {"op": "value", "cell": cell, "sheet_label": True})
            ops.append(_align(cell, "left"))
            next_row = 4
        elif t == "audit_header":
            first = p.get("first_col", 2)
            last = first + 5
            meta = {k: plain_punct(v) for k, v in (p.get("meta") or {}).items()}
            ops.append({"op": "column_width", "col": "A", "width": 2.0})
            box = _rng(first, 2, last, 3)
            for i, labels in enumerate((AUDIT_HEADER_FIELDS[:3], AUDIT_HEADER_FIELDS[3:])):
                r = 2 + i
                for j, lab in enumerate(labels):
                    ops.append({"op": "value", "cell": _cell(first + 2 * j, r), "text": lab})
                    vcell = _cell(first + 2 * j + 1, r)
                    if lab == "시트명" and not meta.get("시트명"):
                        ops.append({"op": "value", "cell": vcell, "sheet_label": True})
                    else:
                        ops.append({"op": "value", "cell": vcell, "text": meta.get(lab, "") or ""})
            ops.append(_font(box, spec, body, False, col["text"]))
            for j in range(3):
                ops.append(_font(_rng(first + 2 * j, 2, first + 2 * j, 3), spec, body, True, col["text"]))
            ops.append(_align(box, "left"))
            ops.append({"op": "border", "range": box, "style": "frame", "sides": "top_bottom",
                        "outer_color": oc, "outer_weight": "medium",
                        "inner_color": ic, "inner_weight": "thin"})
            next_row = 5
        elif t == "section_bar":
            row, first, last = p["row"], p.get("first_col", 2), p.get("last_col", 8)
            bar = _rng(first, row, last, row)
            ops.append({"op": "fill", "range": bar, "color": col["title_bg"]})
            ops.append(_font(bar, spec, body, True, col["title_font"]))
            ops.append(_align(bar, "left"))
            ops.append({"op": "value", "cell": _cell(first, row), "text": plain_punct(p["text"])})
            next_row = row + 2
        elif t == "caption":
            cell = p["cell"]
            text = plain_punct(p["text"])
            if len(text.strip()) > rl["caption_max"]:
                warnings.append("캡션이 너무 김(%d자 > %d자) — 짧은 명사구 제목만 쓰고 설명은 비고 열로."
                                % (len(text.strip()), rl["caption_max"]))
            ops.append({"op": "value", "cell": cell, "text": text})
            ops.append(_font(cell, spec, body, True, _BLACK))
            ops.append(_align(cell, "left"))
            next_row = _split_cell(cell)[1] + 2
        elif t == "tab":
            role = rl["tab_role_aliases"].get(p["role"], p["role"])
            tabs = spec["tab_colors"]
            if role not in tabs:
                raise ValueError("unknown tab role '%s'. available: %s" % (p["role"], list(tabs)))
            ops.append({"op": "tab_color", "color": tabs[role]})
        else:
            raise ValueError("unknown part type: %s" % t)

    return {"theme": spec["name"], "ops": ops, "next_header_row": next_row, "warnings": warnings}
