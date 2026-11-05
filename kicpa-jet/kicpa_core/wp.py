# -*- coding: utf-8 -*-
"""조서 공통 부품 - excel-theme audit 테마 위에서 쓰는 작은 도우미."""
from openpyxl.styles import Alignment, Font
from openpyxl.workbook.defined_name import DefinedName

from excel_theme import get_theme, plain_punct  # noqa: F401  (호출 측에서 같이 씀)

TH = "audit"
ACC = "accounting"
LINK_BLUE = "0563C1"


def font_body():
    t = get_theme(TH)
    return Font(name=t["font_name"], size=t.get("font_size_body", 11))


def put(ws, r, c, text, bold=False, wrap=False):
    """표 밖 안내 문장(테마 본문 글꼴)."""
    cell = ws.cell(r, c, plain_punct(text) if isinstance(text, str) and not text.startswith("=") else text)
    f = font_body()
    cell.font = Font(name=f.name, size=f.size, bold=bold)
    if wrap:
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    return cell


def nm(wb, name, sheet, ref):
    wb.defined_names[name] = DefinedName(name, attr_text="'%s'!%s" % (sheet, ref))


def link(cell, sheet, ref="A1", text=None):
    """같은 통합문서 안 시트로 가는 셀 하이퍼링크(엑셀 하이퍼링크 모양: 파랑 밑줄 · 맑은 고딕 11)."""
    if text is not None:
        cell.value = plain_punct(text)
    cell.hyperlink = "#'%s'!%s" % (sheet, ref)
    f = font_body()
    cell.font = Font(name=f.name, size=f.size, color=LINK_BLUE, underline="single")
    return cell


def merge_text_col(ws, header_row, first_row, last_row, c1, c2, wrap=True):
    """표의 c1~c2 열을 머리글 · 본문 행마다 병합(서식은 write_table 이 모든 칸에 같게 준 상태)."""
    rows = ([header_row] if header_row else []) + list(range(first_row, last_row + 1))
    for r in rows:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        if wrap and r != header_row:
            ws.cell(r, c1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)


def left_top(ws, ref):
    for row in ws[ref]:
        for c in (row if isinstance(row, tuple) else (row,)):
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)


def fmt_date(d):
    return d.strftime("%Y.%m.%d") if d else ""


def write_data_fast(ws, headers, rows, header_row=5, start_col=2, number_cols=None, text_cols=()):
    """줄이 많은 Data_* 시트용 표 쓰기 - write_table 과 같은 모양(회색 머리글 · 삼선표 · 맑은 고딕 11)을
    칸마다 같은 서식 객체로 빠르게 준다(분개장 수십만 줄). number_cols = {0-based 열: 서식키}, text_cols = 왼쪽 정렬 열."""
    from openpyxl.styles import Border, Side
    from excel_theme import NUMBER_FORMATS, style_header_row
    try:
        from excel_theme import _stamp_author
        _stamp_author(ws)
    except ImportError:
        pass
    th = get_theme(TH)
    for j, h in enumerate(headers):
        ws.cell(header_row, start_col + j, h)
    style_header_row(ws, th, header_row, start_col, start_col + len(headers) - 1)
    thin = Side(style="thin", color=th["border_color"])
    thick = Side(style="medium", color=th["border_color_outer"])
    n = len(headers)
    font = Font(name=th["font_name"], size=th["font_size_body"], color=th.get("text_color"))
    al_r = Alignment(horizontal="right", vertical="center")
    al_l = Alignment(horizontal="left", vertical="center")
    fmts = {j: NUMBER_FORMATS.get(v, v) for j, v in (number_cols or {}).items()}
    tset = set(text_cols)

    def border(j, first, last):
        return Border(left=None if j == 0 else thin, right=None if j == n - 1 else thin,
                      top=thick if first else thin, bottom=thick if last else thin)

    hb = [border(j, True, False) for j in range(n)]
    for j in range(n):
        ws.cell(header_row, start_col + j).border = hb[j]
    from copy import copy
    # 열마다(본문 줄 · 마지막 줄) 서식을 한 번만 만들어 두고 칸에는 서식 배열만 복사한다(openpyxl 서식 해시 비용 회피)
    proto = {}
    for last in (False, True):
        for j in range(n):
            pc = ws.cell(header_row, start_col + j)          # 임시로 서식을 만들 칸(머리글 서식은 아래에서 되돌림)
            keep = copy(pc._style)
            pc.font = font
            pc.border = border(j, False, last)
            if j in fmts:
                pc.number_format = fmts[j]
                pc.alignment = al_r
            proto[(j, last, "n")] = copy(pc._style)
            pc.alignment = al_l
            if j in fmts:
                pc.number_format = "General"
            proto[(j, last, "t")] = copy(pc._style)          # 글(왼쪽 · 일반 서식) - 숫자 열에 글이 섞여도 왼쪽
            pc._style = keep
    nr = len(rows)
    for i, row in enumerate(rows):
        r = header_row + 1 + i
        last = i == nr - 1
        for j in range(n):
            v = row[j] if j < len(row) else None
            c = ws.cell(r, start_col + j, v) if v is not None else ws.cell(r, start_col + j)
            if isinstance(v, str):
                k = "t" if (not v.startswith("=") or j in tset) else "n"
            elif v is None:
                k = "t" if j in tset else "n"
            else:
                k = "n"
            c._style = copy(proto[(j, last, k)])
    return header_row + nr


def text_units(s):
    return sum(2 if ord(ch) > 0x1100 else 1 for ch in str(s))


def fit_headers(wb, first_col=2, last_col=7, rows=(2, 3), pad=3):
    """조서 머리(B2:G3) 칸 글이 칸 안에 들어가도록 그 열만 넓힌다(엑셀 배율 차이를 생각해 넉넉히)."""
    from openpyxl.utils import get_column_letter
    for ws in wb.worksheets:
        for c in range(first_col, last_col + 1):
            need = 0
            for r in rows:
                v = ws.cell(r, c).value
                if v is None or (isinstance(v, str) and v.startswith("=")):
                    continue
                need = max(need, text_units(v) * 1.12 + pad)
            col = get_column_letter(c)
            cur = ws.column_dimensions[col].width or 8.43
            if need > cur:
                ws.column_dimensions[col].width = round(need, 1)


def fit_rows(path, timeout=600):
    """엑셀 COM 으로 행높이를 맞춘다 - ① 줄바꿈 칸 행 자동 맞춤(kicpa_core/autofit_rows.ps1, 줄 많은 Data 시트는 건너뜀)
    ② 병합 칸 여러 줄 글(excel-theme tools/fit_merged_rows.ps1). 윈도 · 엑셀이 없으면 건너뜀."""
    import os
    import subprocess
    from . import HERE, THEME_DIR
    steps = [os.path.join(HERE, "autofit_rows.ps1"), os.path.join(THEME_DIR, "tools", "fit_merged_rows.ps1")]
    if os.name != "nt" or not all(os.path.exists(x) for x in steps):
        return "행높이 맞춤 건너뜀(윈도 엑셀 없음) - 엑셀에서 줄바꿈 칸 행높이를 확인"
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets = [n for n in wb.sheetnames if not n.startswith("Data_")]   # 줄 많은 Data 시트는 맞추지 않음
    wb.close()
    msg = []
    for ps1, arg in zip(steps, ("-File", "-Files")):
        q = lambda x: "'%s'" % str(x).replace("'", "''")
        cmd = "& %s %s %s" % (q(ps1), arg, q(path))
        if arg == "-Files":
            cmd += " -Sheets " + ",".join(q(x) for x in sheets)
        try:
            p = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", cmd],
                               capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=timeout)
            out = (p.stdout + p.stderr).strip().splitlines()
            msg.append(out[-1][:60] if out else "완료")
        except subprocess.TimeoutExpired:
            msg.append("시간 초과 - 파일이 다른 엑셀에 열려 있는지 확인")
            break
    return "행높이 맞춤: " + " / ".join(msg)
