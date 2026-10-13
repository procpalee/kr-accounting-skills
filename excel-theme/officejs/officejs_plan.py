# -*- coding: utf-8 -*-
"""
officejs_plan.py — plan.py 의 중립 ops → Office.js 전용 ops + 그대로 실행할 코드.

Claude for Excel(애드인)은 Office.js 로 셀을 서식한다. 중립 ops 를 Claude 가 매번 번역하면
해석 편차가 생기므로, 여기서 번역을 결정론적으로 끝내고 **실행만 하면 되는 코드**를 돌려준다.

- 서식 규칙(무엇을 칠할지)은 plan.py 가 단일 소스 — 이 파일은 표현(Office.js)만 담당한다.
- 원본: excel-theme/officejs/ → tools/build_officejs_skill.py 가 스킬 패키지(scripts/)로 복사한다.
- 정렬 by_type(숫자·수식 오른쪽 / 텍스트 왼쪽) = Excel 가로 정렬 General.
- 열너비: autofitColumns 후 상·하한(문자 단위 → 포인트 환산)으로 자른다(값을 읽어야 해서 sync 1회).
- 긴 대시 치환: 수식이 아닌 텍스트 셀만 값을 다시 쓴다(sync 1회).
"""
import json
import re

from plan import format_plan, sheet_parts_plan

_WEIGHT = {"hairline": "Hairline", "thin": "Thin", "medium": "Medium", "thick": "Thick"}
_ALIGN = {"left": "Left", "center": "Center", "right": "Right", "general": "General",
          "by_type": "General"}
_MAX_DIGIT_PX = 7          # Excel 기본 글꼴 최대 숫자폭(px) — 열너비(문자수) → 포인트 환산용


def _hex(c):
    return "#" + str(c).lstrip("#").upper()


def _col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _dims(address):
    """"C5:E10" → (행 수, 열 수)."""
    m = re.fullmatch(r"([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?", address)
    if not m:
        raise ValueError("unsupported address: %s" % address)
    c1, r1 = m.group(1), int(m.group(2))
    c2, r2 = m.group(3) or c1, int(m.group(4) or r1)
    return r2 - r1 + 1, _col_num(c2) - _col_num(c1) + 1


def _width_points(chars):
    """openpyxl 열너비(문자 단위) → Office.js columnWidth(포인트)."""
    px = int(((256 * chars + int(128 / _MAX_DIGIT_PX)) / 256) * _MAX_DIGIT_PX)
    return round(px * 0.75, 2)


def _edge(address, edge, style="Continuous", weight=None, color=None):
    op = {"op": "border", "address": address, "edge": edge, "style": style}
    if style == "Continuous":
        op["weight"] = _WEIGHT[weight]
    if style != "None":
        op["color"] = _hex(color)
    return op


def _side_edge(address, edge, side):
    s = side["style"]
    if s == "none":
        return _edge(address, edge, "None")
    if s == "double":
        return _edge(address, edge, "Double", color=side["color"])
    return _edge(address, edge, weight=s, color=side["color"])


def _border_ops(o):
    a = o["range"]
    rows, cols = _dims(a)
    style = o["style"]
    oc, ow = o["outer_color"], o.get("outer_weight", "medium")
    ic, iw = o.get("inner_color"), o.get("inner_weight", "thin")
    out = []
    if style == "frame":
        out += [_edge(a, "EdgeTop", weight=ow, color=oc), _edge(a, "EdgeBottom", weight=ow, color=oc),
                _edge(a, "EdgeLeft", "None"), _edge(a, "EdgeRight", "None")]
        if rows > 1:
            out.append(_edge(a, "InsideHorizontal", weight=iw, color=ic))
        if cols > 1:
            out.append(_edge(a, "InsideVertical", weight=iw, color=ic))
    elif style == "grid":
        out += [_edge(a, e, weight=ow, color=oc)
                for e in ("EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight")]
        if rows > 1:
            out.append(_edge(a, "InsideHorizontal", weight=iw, color=ic))
        if cols > 1:
            out.append(_edge(a, "InsideVertical", weight=iw, color=ic))
    elif style == "statement":
        out += [_edge(a, "EdgeTop", weight=ow, color=oc), _edge(a, "EdgeBottom", weight=ow, color=oc)]
    else:
        raise ValueError("unknown border style: %s" % style)
    return out


def to_officejs_ops(neutral_ops):
    """중립 ops → Office.js 속성과 1:1 인 ops (번역 여지 없음)."""
    out = []
    for o in neutral_ops:
        k = o["op"]
        if k == "workbook_font":
            out.append({"op": "workbookFont", "name": o["name"], "size": o["size"]})
        elif k == "unfreeze":
            out.append({"op": "unfreeze"})
        elif k == "column_width":
            out.append({"op": "columnWidth", "address": "%s:%s" % (o["col"], o["col"]),
                        "points": _width_points(o["width"])})
        elif k == "fill":
            out.append({"op": "fill", "address": o["range"], "color": _hex(o["color"])})
        elif k == "font":
            f = {"op": "font", "address": o["range"]}
            for key in ("name", "size", "bold"):
                if key in o:
                    f[key] = o[key]
            if o.get("color"):
                f["color"] = _hex(o["color"])
            out.append(f)
        elif k == "align":
            a = o["range"]
            if o.get("horizontal"):
                out.append({"op": "horizontalAlignment", "address": a, "value": _ALIGN[o["horizontal"]]})
            if o.get("vertical"):
                out.append({"op": "verticalAlignment", "address": a, "value": _ALIGN[o["vertical"]]})
            if "wrap" in o:
                out.append({"op": "wrapText", "address": a, "value": bool(o["wrap"])})
        elif k == "number_format":
            rows, cols = _dims(o["range"])
            out.append({"op": "numberFormat", "address": o["range"], "format": o["format"],
                        "rows": rows, "cols": cols})
            if o.get("align"):
                out.append({"op": "horizontalAlignment", "address": o["range"],
                            "value": _ALIGN[o["align"]]})
                out.append({"op": "verticalAlignment", "address": o["range"], "value": "Center"})
                out.append({"op": "wrapText", "address": o["range"], "value": False})
        elif k == "border":
            out.extend(_border_ops(o))
        elif k == "border_edges":
            if "top" in o:
                out.append(_side_edge(o["range"], "EdgeTop", o["top"]))
            if "bottom" in o:
                out.append(_side_edge(o["range"], "EdgeBottom", o["bottom"]))
        elif k == "border_top":                       # 구 형식 호환
            out.append(_edge(o["range"], "EdgeTop", weight=o.get("weight", "medium"), color=o["color"]))
        elif k == "row_height":                       # 구 형식 호환 — 행높이는 더 이상 만들지 않는다
            continue
        elif k == "conditional_negative":
            out.append({"op": "negativeHighlight", "address": o["range"],
                        "color": _hex(o["color"]), "bold": bool(o.get("bold"))})
        elif k == "autofit_columns":
            _, cols = _dims(o["range"])
            out.append({"op": "autofitColumns", "address": o["range"], "cols": cols,
                        "minPoints": _width_points(o["min_width"]),
                        "maxPoints": _width_points(o["max_width"])})
        elif k == "normalize_punct":
            out.append({"op": "normalizePunct", "address": o["range"]})
        elif k == "value":
            v = {"op": "value", "address": o["cell"]}
            if o.get("sheet_label"):
                v["sheetLabel"] = True
            else:
                v["text"] = o.get("text", "")
            out.append(v)
        elif k == "tab_color":
            out.append({"op": "tabColor", "color": _hex(o["color"]) if o.get("color") else ""})
        else:
            raise ValueError("unknown op: %s" % k)
    return out


def _js(v):
    return json.dumps(v, ensure_ascii=False)


_DASH = "/[\\u2014\\u2013\\u2015]/"


def to_officejs_code(ops, sheet_name="", body_only=False):
    """Office.js ops → Excel.run 코드 문자열(수정 없이 실행).
    body_only=True 면 Excel.run 래퍼 없이 본문만(실행 도구가 이미 context 를 주는 경우)."""
    sheet = ("context.workbook.worksheets.getItem(%s)" % _js(sheet_name) if sheet_name
             else "context.workbook.worksheets.getActiveWorksheet()")
    L = [
        "await Excel.run(async (context) => {",
        "  const ws = %s;" % sheet,
        "  const r = (a) => ws.getRange(a);",
        "  const bd = (a, edge, style, weight, color) => {",
        "    const b = r(a).format.borders.getItem(edge);",
        "    b.style = style;",
        "    if (weight) b.weight = weight;",
        "    if (color) b.color = color;",
        "  };",
        "  const nf = (a, fmt, rows, cols) => {",
        "    r(a).numberFormat = Array.from({ length: rows }, () => Array(cols).fill(fmt));",
        "  };",
        "  const punct = async (a) => {",
        "    const g = r(a); g.load(\"formulas\"); await context.sync();",
        "    g.formulas.forEach((row, i) => row.forEach((v, j) => {",
        "      if (typeof v === \"string\" && !v.startsWith(\"=\") && %s.test(v)) {" % _DASH,
        "        g.getCell(i, j).values = [[v.replace(%sg, \"-\")]];" % _DASH,
        "      }",
        "    }));",
        "  };",
        "  const fit = async (a, n, lo, hi) => {",
        "    const g = r(a); g.format.autofitColumns();",
        "    const cs = []; for (let i = 0; i < n; i++) { const c = g.getColumn(i); c.format.load(\"columnWidth\"); cs.push(c); }",
        "    await context.sync();",
        "    for (const c of cs) c.format.columnWidth = Math.min(Math.max(c.format.columnWidth, lo), hi);",
        "  };",
    ]
    if any(o["op"] == "value" and o.get("sheetLabel") for o in ops):
        L += ["  ws.load(\"name\"); await context.sync();",
              "  const sheetLabel = ws.name.replace(/^\\d+\\s+/, \"\").trim();"]
    for o in ops:
        k = o["op"]
        a = _js(o.get("address"))
        if k == "workbookFont":
            L.append("  { const f = context.workbook.styles.getItem(\"Normal\").font; f.name = %s; f.size = %s; }"
                     % (_js(o["name"]), _js(o["size"])))
        elif k == "unfreeze":
            L.append("  ws.freezePanes.unfreeze();")
        elif k == "columnWidth":
            L.append("  r(%s).format.columnWidth = %s;" % (a, _js(o["points"])))
        elif k == "fill":
            L.append("  r(%s).format.fill.color = %s;" % (a, _js(o["color"])))
        elif k == "font":
            props = ["f.%s = %s;" % (p, _js(o[p])) for p in ("name", "size", "bold", "color") if p in o]
            L.append("  { const f = r(%s).format.font; %s }" % (a, " ".join(props)))
        elif k == "horizontalAlignment":
            L.append("  r(%s).format.horizontalAlignment = %s;" % (a, _js(o["value"])))
        elif k == "verticalAlignment":
            L.append("  r(%s).format.verticalAlignment = %s;" % (a, _js(o["value"])))
        elif k == "wrapText":
            L.append("  r(%s).format.wrapText = %s;" % (a, _js(o["value"])))
        elif k == "numberFormat":
            L.append("  nf(%s, %s, %d, %d);" % (a, _js(o["format"]), o["rows"], o["cols"]))
        elif k == "border":
            L.append("  bd(%s, %s, %s, %s, %s);" % (a, _js(o["edge"]), _js(o["style"]),
                                                   _js(o.get("weight")), _js(o.get("color"))))
        elif k == "negativeHighlight":
            L.append("  { const cf = r(%s).conditionalFormats.add(Excel.ConditionalFormatType.cellValue);"
                     " cf.cellValue.format.font.color = %s; cf.cellValue.format.font.bold = %s;"
                     " cf.cellValue.rule = { formula1: \"0\", operator: \"LessThan\" }; }"
                     % (a, _js(o["color"]), _js(o["bold"])))
        elif k == "autofitColumns":
            L.append("  await fit(%s, %d, %s, %s);" % (a, o["cols"], _js(o["minPoints"]), _js(o["maxPoints"])))
        elif k == "normalizePunct":
            L.append("  await punct(%s);" % a)
        elif k == "value":
            L.append("  r(%s).values = [[%s]];" % (a, "sheetLabel" if o.get("sheetLabel") else _js(o["text"])))
        elif k == "tabColor":
            L.append("  ws.tabColor = %s;" % _js(o["color"]))
        else:
            raise ValueError("unknown officejs op: %s" % k)
    L += ["  await context.sync();", "});"]
    if body_only:
        L = [line[2:] if line.startswith("  ") else line for line in L[1:-1]]
    return "\n".join(L)


_RUN = ("code 를 한 글자도 바꾸지 말고 그대로 실행한다. "
        "행높이는 지정하지 않는다(엑셀 기본값). 틀고정은 해제된다.")


def officejs_plan(spec, sheet_name="", rules=None, body_only=False, **layout):
    """테마 spec + 표 레이아웃 → {theme, ops(Office.js), code, run_instruction}."""
    neutral = format_plan(spec, rules=rules, **layout)
    ops = to_officejs_ops(neutral["ops"])
    return {
        "theme": neutral["theme"],
        "ops": ops,
        "code": to_officejs_code(ops, sheet_name, body_only),
        "run_instruction": _RUN + " 값·수식 입력을 먼저 끝낸 뒤 실행한다(긴 대시 치환·열너비가 값을 읽는다). "
                                  "같은 범위에 다시 실행하면 음수 조건부서식이 중복될 수 있다.",
    }


def officejs_sheet_parts(spec, parts, sheet_name="", rules=None, body_only=False):
    """테마 spec + 시트 요소 목록 → {theme, ops, code, next_header_row, warnings, run_instruction}."""
    neutral = sheet_parts_plan(spec, parts, rules=rules)
    ops = to_officejs_ops(neutral["ops"])
    return {
        "theme": neutral["theme"],
        "ops": ops,
        "code": to_officejs_code(ops, sheet_name, body_only),
        "next_header_row": neutral["next_header_row"],
        "warnings": neutral["warnings"],
        "run_instruction": _RUN + " 표보다 먼저 실행하고, 표는 next_header_row 를 header_row 로 쓴다.",
    }
