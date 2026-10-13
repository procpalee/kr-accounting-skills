# -*- coding: utf-8 -*-
"""theme_lint — 워크북이 excel-theme 테마 규약을 지키는지 점검한다.

사용:  python theme_lint.py <파일.xlsx> [테마이름=default] [시트명 ...]
       (시트명 생략 시 전체 시트 검사, '원본'으로 시작하는 시트는 제외)

점검 항목 (default/audit/procpa 계열):
  1. 폰트     — 테마 font_name(맑은 고딕) 외 폰트, 9pt 미만 크기
  2. 제목     — title_bg(default #393939 / audit #002060) 밴드 + 흰 글씨 + 테마 excel_title_size(11pt)
  3. 채움색   — 테마 토큰(header_bg·note_fill·linked_fill·total_fill·title_bg·header2_bg·
                subheader_fill) 외의 하드코딩 배경색
  4. 표 테두리 — 헤더행(F2F2F2)이 있는데 데이터 영역에 테두리가 전혀 없는 경우
  5. 여백     — A열·1행은 비워 둔다(값 금지). 표·제목은 B2/B열부터
  6. 탭 색     — 테마 tab_colors(역할 5색) 또는 색 없음만 허용. 안내·표지 시트는 색 금지
  7. 제목 밴드  — default 계열(audit 제외)은 B2 제목 밴드에 **시트명만**(앞 순번 뗀 ws.title).
                 설명·부제·회사명·기준일을 덧붙이지 않는다 (사용자 확정 2026-09-10)
  8. 문장부호  — 셀 텍스트·시트명에 긴 대시(— – ―) 금지 → "-" 사용 (가운뎃점 · 은 허용)
                 (사용자 확정 2026-09-11, 모든 테마)
  9. 캡션 길이  — 표 소제목(B열 굵은 글씨, 채움 없음)은 **짧은 제목만**. MAX_CAPTION(40자) 초과 금지
                 설명·근거·판단 사유는 캡션이 아니라 표의 비고 열이나 안내 시트에 쓴다 (사용자 확정 2026-09-12)

종료코드: 위반 0건이면 0, 있으면 1. 산출 직후 이 린트를 돌려 0을 확인한다.
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from openpyxl import load_workbook
from excel_theme import get_theme, sheet_label, _canon, PUNCT_DASHES

_BAD_PUNCT = set(PUNCT_DASHES)
MAX_CAPTION = 40   # 표 소제목(캡션) 최대 글자수 - 설명은 비고 열·안내 시트로 (사용자 확정 2026-09-12)

def norm(rgb):
    s = str(rgb or "")
    return s[-6:].upper() if len(s) >= 6 else s.upper()

def lint(path, theme="default", only=None):
    th = get_theme(theme)
    ok_fills = {norm(th[k]) for k in ("header_bg", "note_fill", "linked_fill", "todo_fill", "total_fill",
                                      "title_bg", "header2_bg", "subheader_fill", "band_fill") if th.get(k)}
    ok_fills |= {"FFFFFF", "00000000", ""}
    font_name = th["font_name"]
    issues = []
    wb = load_workbook(path)
    for ws in wb.worksheets:
        if ws.sheet_state != "visible": continue
        if only and ws.title not in only: continue
        if ws.title.startswith("원본"): continue
        long_caps = []
        bad_font = bad_size = bad_fill = 0
        fill_samples = {}
        punct_cells = []
        for row in ws.iter_rows():
            for c in row:
                if c.value is None and (not c.fill or c.fill.fill_type != "solid"): continue
                if isinstance(c.value, str) and _BAD_PUNCT & set(c.value):
                    punct_cells.append(c.coordinate)
                if (c.column == 2 and isinstance(c.value, str) and c.font and c.font.b
                        and not (c.fill and c.fill.fill_type == "solid"
                                 and norm(c.fill.fgColor.rgb) not in ("FFFFFF", "00000000", ""))
                        and len(c.value.strip()) > MAX_CAPTION):
                    long_caps.append((c.coordinate, len(c.value.strip())))
                f = c.font
                if f and f.name and f.name not in (font_name, "Consolas"):
                    bad_font += 1
                if f and f.size and f.size < 9:
                    bad_size += 1
                if c.fill and c.fill.fill_type == "solid":
                    hexv = norm(c.fill.fgColor.rgb)
                    if hexv and hexv not in ok_fills:
                        bad_fill += 1
                        fill_samples.setdefault(hexv, c.coordinate)
        if bad_font: issues.append(f"[{ws.title}] 테마 외 폰트 {bad_font}셀 (기준: {font_name})")
        if bad_size: issues.append(f"[{ws.title}] 9pt 미만 폰트 {bad_size}셀")
        if bad_fill:
            ex = " · ".join(f"#{k}@{v}" for k, v in list(fill_samples.items())[:4])
            issues.append(f"[{ws.title}] 테마 토큰 외 배경색 {bad_fill}셀 — {ex}")
        # 8. 문장부호 — 긴 대시 대신 "-"
        if _BAD_PUNCT & set(ws.title):
            issues.append(f"[{ws.title}] 시트명 문장부호 위반 — 긴 대시(—)는 - 로 바꾼다")
        if punct_cells:
            issues.append(f"[{ws.title}] 문장부호 위반 {len(punct_cells)}셀 — 긴 대시(—) → '-' "
                          f"(excel_theme.plain_punct) (예: {punct_cells[:5]})")
        if long_caps:
            ex = " · ".join(f"{c}({n}자)" for c, n in long_caps[:4])
            issues.append(f"[{ws.title}] 캡션이 너무 김 {len(long_caps)}셀 (기준 {MAX_CAPTION}자) — "
                          f"제목만 쓰고 설명은 비고 열·안내 시트로: {ex}")
        margin_bad = [c.coordinate for c in ws["A"] if c.value not in (None, "")]
        row1_bad = [c.coordinate for c in ws[1] if c.value not in (None, "")]
        if margin_bad: issues.append(f"[{ws.title}] A열 여백 위반 {len(margin_bad)}셀 (예: {margin_bad[:3]})")
        if row1_bad: issues.append(f"[{ws.title}] 1행 여백 위반 {len(row1_bad)}셀 (예: {row1_bad[:3]})")
        custom_h = [r for r, d in ws.row_dimensions.items() if d.height is not None]
        if custom_h: issues.append(f"[{ws.title}] 행높이 지정 위반 {len(custom_h)}행 — 기본 행높이만 사용 (예: {custom_h[:5]})")
        if ws.freeze_panes not in (None, "A1"):
            issues.append(f"[{ws.title}] 틀고정 금지 위반 — freeze_panes={ws.freeze_panes}")
        # 7. 제목 밴드(B2)는 시트명만 — audit 테마는 조서 헤더를 쓰므로 제외
        if _canon(theme) != "audit":
            b2 = ws["B2"]
            band = (b2.fill and b2.fill.fill_type == "solid"
                    and norm(b2.fill.fgColor.rgb) == norm(th.get("title_bg", "")))
            got = str(b2.value or "").strip()
            want = sheet_label(ws)
            if band and got and got != want:
                issues.append(f"[{ws.title}] 제목 밴드는 시트명만 — B2 '{got}' ≠ 시트명 '{want}'")
    # 6. 탭 색 — 역할 색(또는 색 없음)만. 원본 시트도 검사 대상이라 별도 루프.
    ok_tabs = {norm(v) for v in th.get("tab_colors", {}).values() if v}
    for ws in wb.worksheets:
        if ws.sheet_state != "visible": continue
        if only and ws.title not in only: continue
        tc = ws.sheet_properties.tabColor
        hexv = norm(getattr(tc, "rgb", None)) if tc is not None else ""
        if not hexv: continue
        if hexv not in ok_tabs:
            allowed = " · ".join("#" + c for c in sorted(ok_tabs))
            issues.append(f"[{ws.title}] 탭 색 규약 이탈 #{hexv} — 허용: {allowed}")
        elif ws.title.startswith(("안내", "표지", "가이드")):
            issues.append(f"[{ws.title}] 안내·표지 시트는 탭 색을 주지 않는다 (현재 #{hexv})")

    if issues:
        print("=== theme_lint 위반 (%d건) — 테마: %s ===" % (len(issues), theme))
        for i in issues: print("  " + i)
        return 1
    print("theme_lint OK — 위반 없음 (테마: %s)" % theme)
    return 0

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(2)
    path = sys.argv[1]
    theme = sys.argv[2] if len(sys.argv) > 2 else "default"
    only = set(sys.argv[3:]) or None
    sys.exit(lint(path, theme, only))
