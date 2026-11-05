# -*- coding: utf-8 -*-
"""kicpa-jet - 한공회 표준조서 3600 부정의 검토 18(1) 분개 및 기타 조정사항 테스트 조서(기본 3600D)를 만든다.

  python -X utf8 jet_build.py --journal <파일 | 폴더 | 와일드카드> [...] --company <회사명> --year-end 2026.12.31
        (--tb <xlsx|csv> [--tb-prior <xlsx|csv>] | --tb-dir <xlaudit tb 폴더> --tb-date 2026.06.30 --tb-prior-date 2025.12.31
         | --tb-json <douzone_tb json>)
        [--gaap ifrs|kgaap] [--start 2026-01-01] [--end 2026-06-30] [--pm 819000000] [--author 이름]
        [--wp-no 3600D] [--list-max 100] --out <xlsx> [--overwrite]

시트: 3600D(한공회 3600 18(1) · 19 · 29(2) 요구사항과 결론) · 10 재무보고절차 이해 · 20 모집단 완전성 · 30 선정기준 ·
      40 선정분개 테스트 · Data_전표 · Data_분개장 · Data_TB · Data_키워드(숨김).
계산(합계 · 대사 · 시나리오 판정 · 건수)은 모두 엑셀 수식이고 파라미터는 이름 정의 칸(30 선정기준의 크림색 칸)이다.
산출 옆에 <조서>_summary.json(파이썬 판정 건수 - jet_check.py 가 엑셀 재계산 값과 대조).
"""
import argparse
import datetime as dt
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from kicpa_core import KicpaError, coa as C, journal as J, tb as TBM, wp as W  # noqa: E402  (경로 설정 뒤)
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402
from openpyxl.utils import get_column_letter as CL  # noqa: E402
from excel_theme import write_audit_header, write_section_bar, write_table, mark_cells, plain_punct  # noqa: E402
import jet_rules as R  # noqa: E402

TH, ACC = W.TH, W.ACC
DATA_ROW = 6
LIST_ROWS = 30
MAX_LINES = 1000000
GAAP = {"ifrs": "한국채택국제회계기준", "kgaap": "일반기업회계기준"}

KICPA_3600 = [   # (번호 · 항목, 조서 시트 키, 비고 기본값) - 한공회 2026 3600 부정의 검토 원문(두 판 같음)
    ("18. 경영진에 의한 통제무력화 위험에 대한 감사인의 평가결과에 관계없이, 다음의 감사절차를 계획하고 수행하여야 한다.", None),
    ("⑴ 총계정원장에 기록된 분개와 재무제표 작성시 이루어진 조정사항의 적절성을 테스트. 이러한 테스트를 위한 감사절차를 설계하고 수행할 때 다음의 절차를 수행하여야 한다.", "20"),
    ("1) 분개 및 기타 조정사항의 처리와 관련하여 아래와 같이 부적절하거나 비정상적인 행위가 있었는지에 대하여 재무보고절차에 관여하는 자들에게 질문", "10"),
    ("① 비경상적이거나 적절하지 않은 분개 또는 적절한 증빙 또는 설명없는 분개 또는 기타 조정사항 등의 입력을 요구받거나 인지하고 있는지", "10"),
    ("② 분개 또는 기타의 조정사항 등의 압력을 통한 통제무력화한 경우가 있었는지", "10"),
    ("③ 정상적인 사업과정을 벗어난 경영진에 의한 분개나 기타 조정사항이 기록된 경우가 있었는지", "10"),
    ("2) 보고기간말에 이루어진 분개 및 기타 조정사항을 선택", "30"),
    ("3) 보고기간 전체를 통하여 이루어진 분개 및 기타 조정사항을 테스트할 필요가 있는지 고려", "30"),
    ("19. 식별된 경영진의 통제무력화 위험에 대응하기 위하여 상기 18에 추가하여 다른 감사절차를 수행할 필요가 있는지 여부를 결정하여야 한다. (즉, 18.에 따라 수행하는 절차로는 해결되지 아니하는 경영진의 추가적인 통제무력화 위험이 구체적으로 존재하는 경우)", None),
    ("29. ⑵ 경영진의 통제무력화 위험에 대응하기 위하여 설계된 절차 등 감사절차의 적용결과", "40"),
]


def colmap(keys, start=2):
    return {k: CL(start + i) for i, k in enumerate(keys)}


class Book:
    def __init__(self, wp_no, info):
        self.wp, self.info = wp_no, info
        self.sheets = {"s": wp_no, "10": "10 재무보고절차 이해", "20": "20 모집단 완전성", "30": "30 선정기준",
                       "40": "40 선정분개 테스트", "dv": "Data_전표", "dj": "Data_분개장", "dt": "Data_TB", "dk": "Data_키워드"}

    def meta(self, no):
        i = self.info
        return {"회사명": i["회사명"], "조서번호": "%s-%s" % (self.wp, no) if no else self.wp, "결산일": i["결산일"],
                "조서명": "분개 및 기타 조정사항 테스트", "작성자": i.get("작성자", "")}


def tb_rows_of(tb, lines):
    """TB 행 + 분개장에만 있는 계정 → Data_TB 행(그룹 · 구분 포함)."""
    rows, seen, warns, mapped = [], set(), [], []
    for t in tb:
        k = C.kind_of(t["코드"], t["계정과목"], t.get("구분"))
        rows.append({"코드": t["코드"], "계정과목": t["계정과목"], "구분": k, "그룹": C.jet_group(t["코드"], t["계정과목"], k),
                     "기초": t["기초"] or 0.0, "기말": t["기말"] or 0.0, "출처": "TB", "대사": t["코드"]})
        seen.add(t["코드"])
    names = {}
    for x in lines:
        names.setdefault(x["코드"], x["계정과목"])
    for code in sorted((c for c in names if c not in seen), key=str):
        k = C.kind_of(code, names[code])
        par = parent_of(code, seen)
        rows.append({"코드": code, "계정과목": names[code], "구분": k, "그룹": C.jet_group(code, names[code], k),
                     "기초": 0.0, "기말": 0.0, "출처": "분개장에만 있음", "대사": par if par is not None else code})
        if par is None and C.cls3(code) != 400:          # 400 손익(마감 대체 전용 계정)은 알리지 않음
            warns.append("TB 에 없는 분개장 계정: %s %s (TB 0 으로 대사)" % (code, names[code]))
        else:
            mapped.append("%s→%s" % (code, par))
    if mapped:
        warns.append("분개장 보조 계정을 TB 상위 계정으로 대사(Data_TB '대사 계정' 칸에서 고침): " + ", ".join(mapped))
    for t in rows:
        if t["그룹"] == "미분류":
            warns.append("계정 성격을 정하지 못함(미분류): %s %s - TB 에 구분 열을 넣으면 풀림" % (t["코드"], t["계정과목"]))
    return rows, warns


def parent_of(code, tb_codes):
    """TB 에 없는 분개장 계정 → 끝자리를 0 으로 채운 TB 상위 계정(1030001 → 1030000). 없으면 None."""
    if not isinstance(code, int):
        return None
    s = str(code)
    for k in range(1, min(5, len(s) - 2) + 1):
        p = int(s[:-k] + "0" * k)
        if p != code and p in tb_codes:
            return p
    return None


def re_toggles(rows):
    """이익잉여금 대사 기본값 - TB 대차로 판정.
    전기손익가산 O = 전기말 재무상태표 계정 합이 0 이 아니고 전기 손익 합과 상쇄(TB 가 손익을 따로 둠 → 기초 이익잉여금에 전기 순이익을 더함).
    당기손익포함 O = 기말 재무상태표 계정 합이 0(재무상태표 이익잉여금에 당기 순이익이 들어 있음)."""
    bs = lambda f: sum(r[f] for r in rows if r["그룹"] in ("재무상태표", "이익잉여금"))
    pl = lambda f: sum(r[f] for r in rows if r["그룹"] == "손익")      # 제조원가 계정은 매출원가에 이미 들어 있어 빼고 셈
    tol = 1.0
    prior_add = "O" if abs(bs("기초")) > tol and abs(bs("기초") + pl("기초")) <= tol else "X"
    cur_in = "O" if abs(bs("기말")) <= tol else "X"
    return prior_add, cur_in


def build(lines, vouchers, tb_rows, info, P, judged, out_path, wp_no, list_max):
    B = Book(wp_no, info)
    S = B.sheets
    wb = Workbook()
    wb.remove(wb.active)
    ws = {k: wb.create_sheet(v) for k, v in S.items()}
    nL, nV, nT = len(lines), len(vouchers), len(tb_rows)
    jl, vl, tl = DATA_ROW + nL - 1, DATA_ROW + nV - 1, DATA_ROW + nT - 1
    nm = lambda n, sh, ref: W.nm(wb, n, sh, ref)
    rng = lambda c, a, b: "$%s$%d:$%s$%d" % (c, a, c, b)

    JK = ["순번", "전표순번", "파일", "키", "일자", "번호", "코드", "과목", "차변", "대변", "적요", "거래처코드", "거래처명",
          "구분", "유형", "작성일", "대사", "마감", "원가", "본06", "본07", "추출"]
    JH = ["순번", "전표순번", "파일", "전표키", "일자", "전표번호", "계정코드", "계정과목", "차변", "대변", "적요", "거래처코드",
          "거래처명", "차대 구분", "전표유형", "작성일", "대사 계정", "마감", "원가대체", "JS06 본계정", "JS07 본계정", "추출 시나리오"]
    Jc = colmap(JK)
    VK = ["순번", "키", "일자", "s", "e", "n", "dr", "cr", "diff", "wd", "마감", "원가", "본06", "본07"] + R.SCN + ["추출", "추출여부"]
    VH = ["전표순번", "전표키", "일자", "시작행", "끝행", "줄 수", "차변합계", "대변합계", "차대차이", "작성일", "마감", "원가대체",
          "JS06 본계정", "JS07 본계정"] + R.SCN + ["추출 시나리오", "추출"]
    Vc = colmap(VK)
    TK = ["코드", "과목", "구분", "그룹", "기초", "기말", "줄수", "희소", "제조", "대사", "출처"]
    Tc = colmap(TK)
    SJ, SV, ST = S["dj"], S["dv"], S["dt"]

    for k, c in [("코드", "코드"), ("차변", "차변"), ("대변", "대변"), ("마감", "마감"), ("원가", "원가"), ("파일", "파일"),
                 ("유형", "유형"), ("구분", "구분"), ("키", "전표키"), ("적요", "적요"), ("거래처명", "거래처명"),
                 ("번호", "전표번호"), ("일자", "일자"), ("작성일", "작성일"), ("과목", "계정과목"), ("본06", "본06"), ("본07", "본07"), ("대사", "대사")]:
        nm("분개_" + c, SJ, rng(Jc[k], DATA_ROW, jl))
    for k, c in [("키", "키"), ("일자", "일자"), ("n", "줄수"), ("dr", "차변"), ("diff", "차대차이"), ("마감", "마감"),
                 ("원가", "원가"), ("본06", "본06"), ("본07", "본07"), ("추출", "추출"), ("추출여부", "추출여부")] + [(s, s) for s in R.SCN]:
        nm("전표_" + c, SV, rng(Vc[k], DATA_ROW, vl))
    for k in ("코드", "과목", "그룹", "기초", "기말", "줄수", "희소", "제조", "대사"):
        nm("TB_" + k, ST, rng(Tc[k], DATA_ROW, tl))

    # ── 30 선정기준 (파라미터 이름 정의를 먼저 잡는다) ──
    s30 = ws["30"]
    hdr = write_audit_header(s30, B.meta("30"), theme=TH, tab="input")
    ncol = 1 + len(R.SCN)
    last30 = 2 + ncol - 1
    r = write_section_bar(s30, hdr, "1. 기본 파라미터", theme=TH, last_col=last30)
    base = [("기준일", P["기준일"], "보고기간말(중간 기준이면 분개장 마지막 날) - JS02·JS10 기준", "기준일"),
            ("수행중요성", P["수행중요성"], "2700 중요성 조서 금액" + ("" if info.get("pm_given") else " - 임시값, 확정 후 고칠 것"), "수행중요성"),
            ("마감 대체 전표 제외", P["마감제외"], "O = 손익 집합 계정이 든 마감 대체 전표를 시나리오에서 뺌", "마감제외"),
            ("재공품 계정코드", P["재공품코드"] or None, "원가 대체 전표(재공품 차변 + 제조원가 대변) 식별", "재공품코드"),
            ("기초 이익잉여금에 전기 손익 더함", P["전기손익가산"], "O = TB 가 손익을 따로 두어 전기말 이익잉여금에 전기 순이익이 빠져 있음", "전기손익가산"),
            ("기말 이익잉여금에 당기 손익 포함", P["당기손익포함"], "O = 기말 재무상태표 이익잉여금에 당기 순이익이 들어 있음(마감 후)", "당기손익포함")]
    write_table(s30, ["항목", "값", "설명", "", "", ""], [list(x[:3]) + [None] * 3 for x in base], theme=TH, header_row=r,
                number_cols={1: ACC}, input_cells=["C%d:C%d" % (r + 1, r + len(base))])
    W.merge_text_col(s30, r, r + 1, r + len(base), 4, 7)
    s30["C%d" % (r + 1)].number_format = "yyyy-mm-dd"
    for i, x in enumerate(base):
        nm(x[3], S["30"], "$C$%d" % (r + 1 + i))
        if i in (2, 4, 5):
            s30["C%d" % (r + 1 + i)].number_format = "@"
            s30["C%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="center")
    r = r + len(base) + 2

    r = write_section_bar(s30, r, "2. 시나리오별 파라미터", theme=TH, last_col=last30)
    sp = [("JS01", "최소 전표금액", "=ROUNDDOWN(수행중요성*0.05,-6)", "전표 차변합계가 이 금액 이상인 전표만(0 = 전부)", "최소금액_JS01"),
          ("JS02", "기간(일)", P["JS02_기간"], "0 = 기준일 당일만", "기간_JS02"),
          ("JS02", "기준금액", "=수행중요성", "전표 차변합계 기준", "기준금액_JS02"),
          ("JS03", "끝자리 단위", P["JS03_단위"], "1,000,000 = 백만원 단위로 떨어지는 금액", "단위_JS03"),
          ("JS03", "최소 금액", P["JS03_최소금액"], "이 금액 이상인 줄만", "최소금액_JS03"),
          ("JS06", "본계정 방향", P["JS06_방향"], "차변·대변·양쪽", "방향_JS06"),
          ("JS07", "본계정 방향", P["JS07_방향"], "차변·대변·양쪽", "방향_JS07"),
          ("JS08", "기준 건수", P["JS08_건수"], "기간 중 줄 수가 이 이하인 계정", "건수_JS08"),
          ("JS11", "기준 일수", P["JS11_일수"], "작성일 - 기표일이 이 일수 이상", "일수_JS11")]
    write_table(s30, ["시나리오", "항목", "값", "설명", "", ""], [list(x[:4]) + [None] * 2 for x in sp], theme=TH, header_row=r,
                number_cols={2: ACC}, input_cells=["D%d:D%d" % (r + 1, r + len(sp))])
    W.merge_text_col(s30, r, r + 1, r + len(sp), 5, 7)
    for i, x in enumerate(sp):
        nm(x[4], S["30"], "$D$%d" % (r + 1 + i))
        if isinstance(x[2], str) and not x[2].startswith("="):
            s30["D%d" % (r + 1 + i)].number_format = "@"
            s30["D%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="center")
    r = r + len(sp) + 2

    r = write_section_bar(s30, r, "3. 공휴일·키워드", theme=TH, last_col=last30)
    hk = [[P["공휴일"][i] if i < len(P["공휴일"]) else None, P["키워드"][i] if i < len(P["키워드"]) else None]
          for i in range(LIST_ROWS)]
    write_table(s30, ["공휴일", "적요 키워드"], hk, theme=TH, header_row=r, number_cols={0: "date"},
                input_cells=["B%d:C%d" % (r + 1, r + LIST_ROWS)])
    for i in range(LIST_ROWS):
        s30["C%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="center")
    nm("목록_공휴일", S["30"], "$B$%d:$B$%d" % (r + 1, r + LIST_ROWS))
    nm("목록_키워드", S["30"], "$C$%d:$C$%d" % (r + 1, r + LIST_ROWS))
    r = r + LIST_ROWS + 2

    r = write_section_bar(s30, r, "4. 계정 목록", theme=TH, last_col=last30)
    lists = [("본계정_JS06", "JS06 본계정", "JS06_본계정"), ("허용_JS06", "JS06 허용", "JS06_허용"),
             ("본계정_JS07", "JS07 본계정", "JS07_본계정"), ("허용_JS07", "JS07 허용", "JS07_허용"),
             ("대상_JS09", "JS09 대상", "JS09_대상"), ("마감코드", "마감 대체", "마감코드")]
    heads, rows = [], [[] for _ in range(LIST_ROWS)]
    for n, h, _ in lists:
        heads += [h, "계정과목"]
    for i in range(LIST_ROWS):
        for j, (n, h, key) in enumerate(lists):
            vals = P[key]
            cc = CL(2 + 2 * j)
            rows[i] += [vals[i] if i < len(vals) else None,
                        '=IF(%s%d="","",IFERROR(INDEX(TB_과목,MATCH(%s%d,TB_코드,0)),"없음"))' % (cc, r + 1 + i, cc, r + 1 + i)]
    write_table(s30, heads, rows, theme=TH, header_row=r, number_cols={2 * j: "0" for j in range(len(lists))},
                input_cells=["%s%d:%s%d" % (CL(2 + 2 * j), r + 1, CL(2 + 2 * j), r + LIST_ROWS) for j in range(len(lists))])
    for j, (n, h, _) in enumerate(lists):
        nm(n, S["30"], "$%s$%d:$%s$%d" % (CL(2 + 2 * j), r + 1, CL(2 + 2 * j), r + LIST_ROWS))
        for i in range(LIST_ROWS):
            s30.cell(r + 1 + i, 3 + 2 * j).alignment = Alignment(horizontal="left", vertical="center")
    r = r + LIST_ROWS + 2

    r = write_section_bar(s30, r, "5. 시나리오", theme=TH, last_col=last30)
    labels = ["테스트 구분", "시나리오 명", "한공회 3600 구분", "적용(O/X)", "근거(기준서 240)", "시나리오 선정 사유", "추출 방법",
              "추출 전표 수", "추출 줄 수", "추출 결과", "조서 참조"]
    body = []
    for lab in labels:
        row = [lab]
        for s in R.SCN:
            g, name, basis, how = R.SCN_INFO[s]
            row.append({"테스트 구분": g, "시나리오 명": name,
                        "한공회 3600 구분": "18(1) 2) 보고기간말" if s in R.PERIOD_END else "18(1) 3) 기간 전체",
                        "적용(O/X)": P["적용"][s], "근거(기준서 240)": basis, "추출 방법": how,
                        "추출 전표 수": "=SUM(전표_%s)" % s, "추출 줄 수": "=SUMPRODUCT(전표_줄수,전표_%s)" % s,
                        "조서 참조": "%s-40" % wp_no}.get(lab))
        body.append(row)
    hr = r
    write_table(s30, ["시나리오 ID"] + R.SCN, body, theme=TH, header_row=hr, number_cols={j: ACC for j in range(1, ncol)})
    ri = {lab: hr + 1 + i for i, lab in enumerate(labels)}
    mark_cells(s30, TH, input=["C%d:%s%d" % (ri[lab], CL(last30), ri[lab]) for lab in ("적용(O/X)", "시나리오 선정 사유", "추출 결과")])
    for j, s in enumerate(R.SCN):
        c = CL(3 + j)
        nm("적용_" + s, S["30"], "$%s$%d" % (c, ri["적용(O/X)"]))
        for lab in labels:
            if lab not in ("추출 전표 수", "추출 줄 수"):
                s30["%s%d" % (c, ri[lab])].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    scn_hdr_row = hr
    r = hr + len(labels) + 2
    r = write_section_bar(s30, r, "6. 결론", theme=TH, last_col=last30)
    write_table(s30, ["결론"] + [""] * (ncol - 1), [[None] * ncol], theme=TH, header_row=r, input_cells=["B%d" % (r + 1)])
    W.merge_text_col(s30, r, r + 1, r + 1, 2, last30)
    s30.column_dimensions["B"].width = 37
    for j in range(len(R.SCN)):
        s30.column_dimensions[CL(3 + j)].width = 30

    # ── Data_키워드(숨김) - 키워드 목록을 가로로(전표 판정 SUMPRODUCT 용) ──
    sk = ws["dk"]
    write_audit_header(sk, B.meta(None), theme=TH, tab="calc")
    write_table(sk, ["키워드 %d" % (i + 1) for i in range(LIST_ROWS)],
                [['=IF(INDEX(목록_키워드,%d)="","",INDEX(목록_키워드,%d))' % (i + 1, i + 1) for i in range(LIST_ROWS)]],
                theme=TH, header_row=5)
    for i in range(LIST_ROWS):
        sk.cell(6, 2 + i).alignment = Alignment(horizontal="left", vertical="center")
    nm("목록_키워드행", S["dk"], "$B$6:$%s$6" % CL(1 + LIST_ROWS))
    sk.sheet_state = "hidden"

    # ── Data_TB ──
    st = ws["dt"]
    write_audit_header(st, B.meta(None), theme=TH, tab="pbc")
    trow, rows = {}, []
    for i, t in enumerate(tb_rows):
        rr = DATA_ROW + i
        trow[t["코드"]] = rr
        c = lambda k: "%s%d" % (Tc[k], rr)
        rows.append([t["코드"], t["계정과목"], t["구분"], t["그룹"], t["기초"], t["기말"],
                     "=COUNTIF(분개_코드,%s)" % c("코드"), "=--AND(%s>0,%s<=건수_JS08)" % (c("줄수"), c("줄수")),
                     '=--(%s="제조")' % c("그룹"), t["대사"], t["출처"]])
    write_table(st, ["계정코드", "계정과목", "구분", "그룹", "기초(전기말)", "기말", "분개장 줄 수", "드문 계정", "제조원가", "대사 계정",
                     "출처"], rows, theme=TH, header_row=5, number_cols={0: "0", 4: ACC, 5: ACC, 6: ACC, 7: "0", 8: "0", 9: "0"},
                input_cells=["%s%d:%s%d" % (Tc["대사"], DATA_ROW, Tc["대사"], tl)])

    # ── Data_분개장 ──
    sj = ws["dj"]
    write_audit_header(sj, B.meta(None), theme=TH, tab="pbc")
    vidx = {}
    for n, v in enumerate(vouchers):
        for k in range(v["s"], v["e"] + 1):
            vidx[k] = n + 1
    rows = []
    vref = lambda key, rr: "=INDEX(전표_%s,%s%d)" % (key, Jc["전표순번"], rr)
    for k, x in enumerate(lines):
        rr = DATA_ROW + k
        rows.append([k + 1, vidx[k], x["파일"], x["키"], x["일자"], x["전표번호"], x["코드"], x["계정과목"], x["차변"], x["대변"],
                     x["적요"] or None, x["거래처코드"] or None, x["거래처명"] or None, x["구분"] or None, x["유형"] or None,
                     x["작성일"], "=IFERROR(INDEX(TB_대사,MATCH(%s%d,TB_코드,0)),%s%d)" % (Jc["코드"], rr, Jc["코드"], rr), vref("마감", rr), vref("원가", rr), vref("본06", rr), vref("본07", rr), vref("추출", rr)])
    W.write_data_fast(sj, JH, rows, header_row=5, number_cols={0: "0", 1: "0", 4: "date", 6: "0", 8: ACC, 9: ACC, 15: "date", 16: "0", 17: "0",
                                                                18: "0", 19: "0", 20: "0"},
                      text_cols=(2, 3, 5, 7, 10, 11, 12, 13, 14, 21))
    for c, w in zip(JK, (9, 9, 22, 18, 12, 9, 11, 22, 16, 16, 30, 11, 22, 9, 9, 12, 11, 7, 9, 9, 9, 18)):
        sj.column_dimensions[Jc[c]].width = w

    # ── Data_전표 ──
    sv = ws["dv"]
    write_audit_header(sv, B.meta(None), theme=TH, tab="calc")
    rows = []
    for n, v in enumerate(vouchers):
        rr = DATA_ROW + n
        c = lambda key: "$%s%d" % (Vc[key], rr)
        blk = lambda key: "INDEX('%s'!$%s:$%s,%s):INDEX('%s'!$%s:$%s,%s)" % (SJ, Jc[key], Jc[key], c("s"), SJ, Jc[key], Jc[key], c("e"))
        gate = lambda s, cond: '=IF(OR(적용_%s<>"O",AND(마감제외="O",%s=1)),0,%s)' % (s, c("마감"), cond)
        amt = "(ABS(%s)+ABS(%s))" % (blk("차변"), blk("대변"))
        side = lambda s: '((방향_%s="양쪽")+(방향_%s="대변")*(%s<>0)+(방향_%s="차변")*(%s<>0))' % (s, s, blk("대변"), s, blk("차변"))
        tags = "&".join('IF(%s=1,", %s","")' % (c(s), s) for s in R.SCN)
        rows.append([
            n + 1, v["키"], v["일자"], DATA_ROW + v["s"], DATA_ROW + v["e"],
            "=%s-%s+1" % (c("e"), c("s")), "=SUM(%s)" % blk("차변"), "=SUM(%s)" % blk("대변"),
            "=%s-%s" % (c("dr"), c("cr")), "=MAX(%s)" % blk("작성일"),
            "=--(SUMPRODUCT(COUNTIF(마감코드,%s))>0)" % blk("코드"),
            '=--AND(재공품코드<>"",COUNTIF(%s,재공품코드)>0,SUMPRODUCT(SUMIFS(TB_제조,TB_코드,%s)*(%s<>0))>0)'
            % (blk("코드"), blk("코드"), blk("대변")),
            "=--(SUMPRODUCT(COUNTIF(본계정_JS06,%s)*%s)>0)" % (blk("코드"), side("JS06")),
            "=--(SUMPRODUCT(COUNTIF(본계정_JS07,%s)*%s)>0)" % (blk("코드"), side("JS07")),
            gate("JS01", "--AND(OR(WEEKDAY(%s,2)>5,COUNTIF(목록_공휴일,%s)>0),%s>=최소금액_JS01)" % (c("일자"), c("일자"), c("dr"))),
            gate("JS02", "--AND(%s>=기준일-기간_JS02,%s<=기준일,%s>=기준금액_JS02)" % (c("일자"), c("일자"), c("dr"))),
            gate("JS03", "--(SUMPRODUCT((%s>=최소금액_JS03)*(MOD(%s,단위_JS03)=0))>0)" % (amt, amt)),
            gate("JS04", '--(SUMPRODUCT((목록_키워드행<>"")*ISNUMBER(SEARCH(목록_키워드행,%s)))>0)' % blk("적요")),
            gate("JS05", "--(SUMPRODUCT(--(LEN(TRIM(%s))=0))>0)" % blk("적요")),
            gate("JS06", "--AND(%s=1,SUMPRODUCT((COUNTIF(본계정_JS06,%s)=0)*(COUNTIF(허용_JS06,%s)=0))>0)" % (c("본06"), blk("코드"), blk("코드"))),
            gate("JS07", "--AND(%s=1,SUMPRODUCT((COUNTIF(본계정_JS07,%s)=0)*(COUNTIF(허용_JS07,%s)=0))>0)" % (c("본07"), blk("코드"), blk("코드"))),
            gate("JS08", "--(SUMPRODUCT(SUMIFS(TB_희소,TB_코드,%s))>0)" % blk("코드")),
            gate("JS09", "--(SUMPRODUCT(COUNTIF(대상_JS09,%s)*(LEN(TRIM(%s))=0))>0)" % (blk("코드"), blk("거래처명"))),
            gate("JS10", "--(%s>기준일)" % c("wd")),
            gate("JS11", "--AND(%s>0,%s-%s>=일수_JS11)" % (c("wd"), c("wd"), c("일자"))),
            "=MID(%s,3,200)" % tags, "=--(LEN(%s)>0)" % c("추출"),
        ])
    nf = {0: "0", 2: "date", 3: "0", 4: "0", 5: "0", 6: ACC, 7: ACC, 8: ACC, 9: "yyyy-mm-dd;;"}
    nf.update({j: "0" for j in range(10, 14 + len(R.SCN))})
    W.write_data_fast(sv, VH, rows, header_row=5, number_cols=nf, text_cols=(1, 14 + len(R.SCN)))
    sv.column_dimensions[Vc["키"]].width = 18
    sv.column_dimensions[Vc["추출"]].width = 22

    # ── 20 모집단 완전성 ──
    s20 = ws["20"]
    hdr = write_audit_header(s20, B.meta("20"), theme=TH, tab="output")
    last20 = 11
    r = write_section_bar(s20, hdr, "1. 목적", theme=TH, last_col=last20)
    W.put(s20, r - 1, 2, "총계정원장에 기록된 분개 전체(모집단)를 입수하여 완전성을 확인한 뒤 시나리오 선정에 쓴다(한공회 3600 18(1)).")
    r += 1
    r = write_section_bar(s20, r, "2. 기간", theme=TH, last_col=last20)
    W.put(s20, r - 1, 2, "기표일 기준 %s ~ %s 에 기록된 모든 분개를 대상으로 한다." % (info["시작"].isoformat(), info["끝"].isoformat()))
    r += 1
    r = write_section_bar(s20, r, "3. 입수 데이터", theme=TH, last_col=last20)
    fl = info["파일목록"]
    rec = []
    for i, f in enumerate(fl):
        rr = r + 1 + i
        tot = f["원본합계"]
        rec.append([f["파일"], f["형식"], "=COUNTIF(분개_파일,B%d)" % rr, "=SUMIFS(분개_차변,분개_파일,B%d)" % rr,
                    "=SUMIFS(분개_대변,분개_파일,B%d)" % rr, tot[0] if tot else None, tot[1] if tot else None,
                    '=IF(G%d="","",E%d-G%d)' % (rr, rr, rr), '=IF(H%d="","",F%d-H%d)' % (rr, rr, rr)])
    f0, f1 = r + 1, r + len(fl)
    rec.append(["합계", None] + ["=SUM(%s%d:%s%d)" % (CL(4 + j), f0, CL(4 + j), f1) for j in range(7)])
    write_table(s20, ["파일", "형식", "줄 수", "차변 합계", "대변 합계", "원본 합계 줄(차변)", "원본 합계 줄(대변)", "차이(차변)", "차이(대변)"],
                rec, theme=TH, header_row=r, number_cols={j: ACC for j in range(2, 9)}, total_rows=[len(rec)])
    rec_tot = f1 + 1
    r = rec_tot + 2
    write_table(s20, ["입수일", "제공자", "입수 경로"], [[None, None, None]], theme=TH, header_row=r, input_cells=["B%d:D%d" % (r + 1, r + 1)],
                secondary=True)
    r += 3
    r = write_section_bar(s20, r, "4. 데이터 필드", theme=TH, last_col=last20)
    fields = [("전표번호", "번호", "전표번호", "COUNTA(분개_전표번호)"), ("계정코드", "코드", "계정코드", "COUNTA(분개_코드)"),
              ("계정과목", "과목", "계정과목", "COUNTA(분개_과목)"), ("기표일", "일자", "일자", "COUNT(분개_일자)"),
              ("차변금액", "차변", "차변", "COUNT(분개_차변)"), ("대변금액", "대변", "대변", "COUNT(분개_대변)"),
              ("적요", "적요", "적요", "COUNTA(분개_적요)"), ("거래처", "거래처명", "거래처명", "COUNTA(분개_거래처명)"),
              ("전표유형", "유형", "유형", "COUNTA(분개_유형)"), ("차대 구분", "구분", "구분", "COUNTA(분개_구분)"),
              ("작성일(입력일)", "작성일", "작성일", "COUNT(분개_작성일)"), ("입력자·승인자", None, "입력자", None)]
    frows = []
    for lab, key, fld, f in fields:
        if key is None or not info["열"].get(fld):
            frows.append([lab, "없음(분개장에 없는 열)", None])
        else:
            frows.append([lab, JH[JK.index(key)], "=" + f])
    write_table(s20, ["요청 필드", "분개장 열", "값이 있는 줄 수"], frows, theme=TH, header_row=r, number_cols={2: ACC})
    r = r + len(frows) + 2

    r = write_section_bar(s20, r, "5. 완전성 검토", theme=TH, last_col=last20)
    head = ["계정코드", "계정과목", "그룹", "기초(TB)", "분개장 차변", "분개장 대변", "계산 기말", "기말(TB)", "차이", "차이 사유"]
    closing = {k for k, j in judged.items() if j["마감"]}
    tmap = {t["코드"]: t["대사"] for t in tb_rows}
    moved_to = {tmap.get(x["코드"], x["코드"]) for x in lines if x["키"] not in closing and (x["차변"] or x["대변"])}
    order = [("재무상태표", "재무상태표 계정"), ("미분류", "미분류 계정"), ("이익잉여금합계", "이익잉여금"),
             ("손익", "손익 계정"), ("제조", "제조원가 계정")]
    body, sect = [], []
    fr = r + 1
    re_sum = lambda col, g: "SUMPRODUCT(SUMIFS(분개_%s,분개_대사,TB_코드,분개_마감,0)*(TB_그룹=\"%s\"))" % (col, g)
    for g, title in order:
        if g == "이익잉여금합계":
            items = [None] if any(t["그룹"] == "이익잉여금" for t in tb_rows) else []
        else:
            items = [t for t in tb_rows if t["그룹"] == g and t["대사"] == t["코드"]
                     and (t["기초"] or t["기말"] or t["코드"] in moved_to)]
        if not items:
            continue
        body.append([title] + [None] * (len(head) - 1))
        sect.append(len(body))
        for t in items:
            rr = fr + len(body)
            if t is None:
                op = '=SUMIF(TB_그룹,"이익잉여금",TB_기초)+IF(전기손익가산="O",SUMIF(TB_그룹,"손익",TB_기초),0)'
                dr = '=%s+IF(당기손익포함="O",%s+%s,0)' % (re_sum("차변", "이익잉여금"), re_sum("차변", "손익"), re_sum("차변", "제조"))
                cr = dr.replace("차변", "대변")
                body.append(["이익잉여금", "이익잉여금 계정 합계", "이익잉여금", op, dr, cr, "=E%d+F%d-G%d" % (rr, rr, rr),
                             '=SUMIF(TB_그룹,"이익잉여금",TB_기말)', "=I%d-H%d" % (rr, rr), None])
                continue
            tb = trow[t["코드"]]
            ref = lambda col: "='%s'!%s%d" % (ST, Tc[col], tb)
            if g == "제조":
                dr = "=SUMIFS(분개_차변,분개_대사,B%d,분개_마감,0,분개_원가,0)" % rr
            else:
                dr = "=SUMIFS(분개_차변,분개_대사,B%d,분개_마감,0)" % rr
            cr = dr.replace("분개_차변", "분개_대변")
            body.append([ref("코드"), ref("과목"), ref("그룹"), ref("기초") if g in ("재무상태표", "미분류") else 0, dr, cr,
                         "=E%d+F%d-G%d" % (rr, rr, rr), ref("기말"), "=I%d-H%d" % (rr, rr), None])
    write_table(s20, head, body, theme=TH, header_row=r, section_rows=sect,
                number_cols={0: "0", 3: ACC, 4: ACC, 5: ACC, 6: ACC, 7: ACC, 8: ACC},
                input_cells=["K%d:K%d" % (fr, fr + len(body) - 1)])
    for i in range(len(body)):
        s20["K%d" % (fr + i)].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        s20["C%d" % (fr + i)].alignment = Alignment(horizontal="left", vertical="center")
    comp_first, comp_last = fr, fr + len(body) - 1
    r = comp_last + 1
    for t in ["금액은 차변 +, 대변 -. 마감 대체 전표는 빼고, 제조원가 계정은 원가 대체 전표도 뺀다. 손익 계정 기초는 0.",
              "TB 에 없는 분개장 보조 계정은 Data_TB '대사 계정'의 TB 계정으로 모아 대사한다(1030001 → 1030000).",
              "이익잉여금 줄 = 이익잉여금 계정 합계. 30 선정기준 1. 기본 파라미터의 두 칸으로 전기·당기 손익을 더할지 정한다."]:
        W.put(s20, r, 2, t)
        r += 1
    r += 1
    r = write_section_bar(s20, r, "6. 검증 요약", theme=TH, last_col=last20)
    I = "$J$%d:$J$%d" % (comp_first, comp_last)
    K = "$K$%d:$K$%d" % (comp_first, comp_last)
    summ = [["분개장 차변 - 대변", "=SUM(분개_차변)-SUM(분개_대변)", "0 이어야 함"],
            ["차대가 다른 전표 수", "=SUMPRODUCT(--(ABS(전표_차대차이)>0.5))", "0 이어야 함"],
            ["원본 합계 줄과 차이", "=SUM(I%d:J%d)" % (rec_tot, rec_tot), "파일별 '합계' 줄 대사"],
            ["차이가 있는 계정 수", "=SUMPRODUCT(--(ABS(%s)>0.5))" % I, "차이 사유를 적는다"],
            ["차이 합계", "=SUM(%s)" % I, "재분류·상계처럼 짝을 이루는 차이면 0 에 가까움"],
            ["차이 사유가 빈 계정 수", '=SUMPRODUCT((ABS(%s)>0.5)*(%s=""))' % (I, K), "0 이 되도록"]]
    write_table(s20, ["항목", "값", "기준", "", "", "", ""], [x + [None] * 4 for x in summ], theme=TH, header_row=r, number_cols={1: ACC})
    W.merge_text_col(s20, r, r + 1, r + len(summ), 4, 8)
    sum20 = {s[0]: r + 1 + i for i, s in enumerate(summ)}
    for c, w in (("B", 32), ("C", 34), ("D", 16), ("E", 18), ("F", 18), ("G", 20), ("H", 20), ("I", 18), ("J", 18), ("K", 30)):
        s20.column_dimensions[c].width = w

    # ── 40 선정분개 테스트 ──
    s40 = ws["40"]
    hdr = write_audit_header(s40, B.meta("40"), theme=TH, tab="output")
    r = hdr
    by = {s: sorted((v for v in vouchers if judged[v["키"]][s]), key=lambda v: -judged[v["키"]]["차변합계"]) for s in R.SCN}
    tbname = {t["코드"]: t["계정과목"] for t in tb_rows}
    for no, s in enumerate(R.SCN, 1):
        r = write_section_bar(s40, r, "%d) %s %s" % (no, s, R.SCN_INFO[s][1]), theme=TH, last_col=12)
        if s in ("JS06", "JS07") and P["적용"][s] == "O":
            vk = {v["키"] for v in vouchers if judged[v["키"]]["본" + s[2:]]}
            codes = sorted({x["코드"] for x in lines if x["키"] in vk}, key=str)
            if codes:
                rows = []
                for cd in codes:
                    rr = r + 1 + len(rows)
                    rows.append([cd, tbname.get(cd, ""),
                                 "=SUMIFS(분개_차변,분개_코드,B%d,분개_본%s,1)" % (rr, s[2:]),
                                 "=SUMIFS(분개_대변,분개_코드,B%d,분개_본%s,1)" % (rr, s[2:]),
                                 '=IF(COUNTIF(본계정_%s,B%d)>0,"본계정",IF(COUNTIF(허용_%s,B%d)>0,"허용 상대계정","허용 외"))' % (s, rr, s, rr)])
                n0 = r + 1
                rows.append(["합계", None, "=SUM(D%d:D%d)" % (n0, n0 + len(codes) - 1), "=SUM(E%d:E%d)" % (n0, n0 + len(codes) - 1), None])
                W.put(s40, r - 1, 2, "상대계정 분석 - 본계정이 든 전표의 계정별 금액")
                write_table(s40, ["계정코드", "계정과목", "차변", "대변", "상대계정 구분"], rows, theme=TH, header_row=r,
                            number_cols={0: "0", 2: ACC, 3: ACC}, total_rows=[len(rows)], secondary=True)
                for i in range(len(rows)):
                    s40["F%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="center")
                r = r + len(rows) + 2
        vs = by[s]
        if P["적용"][s] != "O":
            W.put(s40, r - 1, 2, "적용하지 않음(30 선정기준 적용 X)")
            r += 1
            continue
        if not vs:
            W.put(s40, r - 1, 2, "추출된 전표 없음")
            r += 1
            continue
        if len(vs) > list_max:
            W.put(s40, r - 1, 2, "추출 %s건 중 차변합계가 큰 %d건 - 나머지는 Data_전표 '추출 시나리오' 열로 거른다"
                  % ("{:,}".format(len(vs)), list_max))
            r += 1
            vs = vs[:list_max]
        rows = []
        for i, v in enumerate(vs):
            rr = r + 1 + i
            L = lines[v["s"]:v["e"] + 1]
            def uniq(side):                          # 계정 3개까지, 나머지는 '외 n개'(한 칸 3줄 이내)
                u = list(dict.fromkeys(x["계정과목"] for x in L if x[side]))
                return ", ".join(u[:3]) + (" 외 %d개" % (len(u) - 3) if len(u) > 3 else "")
            memo = next((x["적요"] for x in L if x["적요"]), "")
            cust = next((x["거래처명"] for x in L if x["거래처명"]), "")
            rows.append([i + 1, v["키"], v["일자"], "=INDEX(전표_차변,MATCH(C%d,전표_키,0))" % rr,
                         uniq("차변") or None, uniq("대변") or None, memo or None, cust or None, None, None, None])
        write_table(s40, ["No", "전표키", "일자", "차변합계", "차변 계정", "대변 계정", "적요", "거래처", "증빙 확인", "사업상 근거",
                          "검토 결과"], rows, theme=TH, header_row=r, number_cols={0: "0", 2: "date", 3: ACC},
                    input_cells=["J%d:L%d" % (r + 1, r + len(rows))])
        for i in range(len(rows)):
            for col in "CFGHIJKL":
                s40["%s%d" % (col, r + 1 + i)].alignment = Alignment(horizontal="left", vertical="center",
                                                                    wrap_text=col in "FGHIJKL")
        r = r + len(rows) + 2
    for c, w in (("B", 8), ("C", 21), ("D", 18), ("E", 18), ("F", 44), ("G", 44), ("H", 48), ("I", 32), ("J", 16), ("K", 20), ("L", 24)):
        s40.column_dimensions[c].width = w

    # ── 10 재무보고절차 이해 ──
    s10 = ws["10"]
    hdr = write_audit_header(s10, B.meta("10"), theme=TH, tab="input")
    r = write_section_bar(s10, hdr, "1. 담당자 인터뷰", theme=TH, last_col=8)
    write_table(s10, ["담당자명", "담당부서", "직책", "인터뷰 일시", "인터뷰 방법"], [[None] * 5 for _ in range(3)], theme=TH,
                header_row=r, input_cells=["B%d:F%d" % (r + 1, r + 3)])
    r += 5
    r = write_section_bar(s10, r, "2. 프로세스 이해", theme=TH, last_col=8)
    items = [("분개", "분개 생성(자동 전표·수기 전표)"), ("분개", "분개 승인"), ("분개", "입력·승인 권한 관리"),
             ("결산", "결산 조정 분개 절차"), ("결산", "재무제표 작성 시 조정사항")]
    write_table(s10, ["구분", "항목", "내용", "", "", "", ""], [[a, b] + [None] * 5 for a, b in items], theme=TH,
                header_row=r, input_cells=["D%d:H%d" % (r + 1, r + len(items))])
    W.merge_text_col(s10, r, r + 1, r + len(items), 4, 8)
    r = r + len(items) + 2
    r = write_section_bar(s10, r, "3. 관여자 질문", theme=TH, last_col=8)
    qs = [x[0] for x in KICPA_3600[3:6]]
    write_table(s10, ["질문(3600 18(1) 1))", "", "", "질문 대상", "일자", "답변", "조서 참조"],
                [[q, None, None, None, None, None, "3600C"] for q in qs], theme=TH, header_row=r,
                input_cells=["E%d:G%d" % (r + 1, r + len(qs))])
    W.merge_text_col(s10, r, r + 1, r + len(qs), 2, 4)
    for i in range(len(qs)):
        s10["G%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    r = r + len(qs) + 2
    r = write_section_bar(s10, r, "4. 월별 분개 분포", theme=TH, last_col=8)
    months = []
    d0 = dt.date(info["시작"].year, info["시작"].month, 1)
    while d0 <= info["끝"]:
        months.append(d0)
        d0 = dt.date(d0.year + (d0.month == 12), d0.month % 12 + 1, 1)
    rows = []
    for m in months:
        rr = r + 1 + len(rows)
        rows.append([m, '=COUNTIFS(전표_일자,">="&B%d,전표_일자,"<"&EDATE(B%d,1))' % (rr, rr),
                     '=COUNTIFS(분개_일자,">="&B%d,분개_일자,"<"&EDATE(B%d,1))' % (rr, rr),
                     '=SUMIFS(분개_차변,분개_일자,">="&B%d,분개_일자,"<"&EDATE(B%d,1))' % (rr, rr)])
    n0 = r + 1
    rows.append(["합계"] + ["=SUM(%s%d:%s%d)" % (CL(3 + j), n0, CL(3 + j), n0 + len(months) - 1) for j in range(3)])
    write_table(s10, ["월", "전표 수", "줄 수", "차변 합계"], rows, theme=TH, header_row=r,
                number_cols={0: "yyyy-mm", 1: ACC, 2: ACC, 3: ACC}, total_rows=[len(rows)])
    for i in range(len(rows)):                       # COUNTIFS 조건의 & 때문에 글로 판정되지 않게
        for col in "CDE":
            s10["%s%d" % (col, r + 1 + i)].alignment = Alignment(horizontal="right", vertical="center")
    r = r + len(rows) + 2
    if info["열"].get("유형"):
        r = write_section_bar(s10, r, "5. 전표 유형 분포", theme=TH, last_col=8)
        kinds = sorted({x["유형"] for x in lines if x["유형"]})
        rows = []
        for k in kinds:
            rr = r + 1 + len(rows)
            rows.append([k, "=COUNTIF(분개_유형,B%d)" % rr, "=SUMIFS(분개_차변,분개_유형,B%d)" % rr])
        n0 = r + 1
        rows.append(["합계", "=SUM(C%d:C%d)" % (n0, n0 + len(kinds) - 1), "=SUM(D%d:D%d)" % (n0, n0 + len(kinds) - 1)])
        write_table(s10, ["전표 유형", "줄 수", "차변 합계"], rows, theme=TH, header_row=r, number_cols={1: ACC, 2: ACC},
                    total_rows=[len(rows)])
        r = r + len(rows) + 2
    r = write_section_bar(s10, r, "%d. 결론" % (6 if info["열"].get("유형") else 5), theme=TH, last_col=8)
    write_table(s10, ["결론", "", "", "", "", "", ""], [[None] * 7], theme=TH, header_row=r, input_cells=["B%d" % (r + 1)])
    W.merge_text_col(s10, r, r + 1, r + 1, 2, 8)
    for c, w in (("B", 14), ("C", 36), ("D", 16), ("E", 16), ("F", 14), ("G", 30), ("H", 12)):
        s10.column_dimensions[c].width = w

    # ── 요약(3600D) ──
    s0 = ws["s"]
    hdr = write_audit_header(s0, B.meta(None), sheet_title="분개 및 기타 조정사항 테스트", theme=TH, tab="output")
    r = write_section_bar(s0, hdr, "1. 기본 정보", theme=TH, last_col=8)
    basic = [["회계기준", GAAP[info["gaap"]]], ["대상 기간", "%s ~ %s" % (info["시작"].isoformat(), info["끝"].isoformat())],
             ["분개장", "%s(%s)" % (info["파일"], info["형식"])], ["TB", info["tb"]],
             ["수행중요성", "=수행중요성"], ["관련 조서", "3600 부정의 검토 18(1), 3600C 재무보고절차 관여자에 대한 질문서"]]
    write_table(s0, ["항목", "내용", "", "", "", "", ""], [x + [None] * 5 for x in basic], theme=TH, header_row=r, number_cols={})
    W.merge_text_col(s0, r, r + 1, r + len(basic), 3, 8)
    s0["C%d" % (r + 5)].number_format = "#,###,##0;[Red](#,###,##0);-"
    s0["C%d" % (r + 5)].alignment = Alignment(horizontal="right", vertical="center")
    r = r + len(basic) + 2
    r = write_section_bar(s0, r, "2. 한공회 3600 요구사항", theme=TH, last_col=8)
    target = {"10": S["10"], "20": S["20"], "30": S["30"], "40": S["40"]}
    rows = [[plain_punct(t), None, None, None, "%s-%s" % (wp_no, k) if k else None, None, None] for t, k in KICPA_3600]
    write_table(s0, ["항목", "", "", "", "조서번호", "수행자", "비고"], rows, theme=TH, header_row=r,
                input_cells=["G%d:H%d" % (r + 1, r + len(rows))])
    W.merge_text_col(s0, r, r + 1, r + len(rows), 2, 5)
    for i, (t, k) in enumerate(KICPA_3600):
        cell = s0["F%d" % (r + 1 + i)]
        if k:
            W.link(cell, target[k])
        s0["H%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    r = r + len(rows) + 2
    r = write_section_bar(s0, r, "3. 테스트 요약", theme=TH, last_col=8)
    tsum = [["모집단 줄 수", "=COUNT(분개_차변)"], ["모집단 전표 수", "=COUNT(전표_줄수)"],
            ["완전성 차이 계정 수", "='%s'!C%d" % (S["20"], sum20["차이가 있는 계정 수"])],
            ["적용한 시나리오 수", '=COUNTIF(%s!$C$%d:$%s$%d,"O")' % ("'%s'" % S["30"], ri["적용(O/X)"], CL(last30), ri["적용(O/X)"])],
            ["추출 전표 수(합집합)", "=SUM(전표_추출여부)"],
            ["보고기간말 시나리오 추출 전표 수", "=SUMPRODUCT(--((%s)>0))" % "+".join("전표_%s" % s for s in R.PERIOD_END)],
            ["예외 사항 수", None]]
    write_table(s0, ["항목", "값"], tsum, theme=TH, header_row=r, number_cols={1: ACC}, input_cells=["C%d" % (r + len(tsum))])
    r = r + len(tsum) + 2
    r = write_section_bar(s0, r, "4. 결론", theme=TH, last_col=8)
    write_table(s0, ["결론", "", "", "", "", "", ""], [[None] * 7], theme=TH, header_row=r, input_cells=["B%d" % (r + 1)])
    W.merge_text_col(s0, r, r + 1, r + 1, 2, 8)
    for c, w in (("B", 37), ("C", 30), ("D", 16), ("E", 16), ("F", 14), ("G", 12), ("H", 30)):
        s0.column_dimensions[c].width = w

    W.fit_headers(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return {"시나리오 머리글 행": scn_hdr_row, "20 검증 요약": sum20}


def completeness(lines, tb_rows, judged, prior_add, cur_in):
    """20 완전성 검토와 같은 계산(파이썬) → {계정코드: 차이 = 기말(TB) - 계산 기말} (0.5 넘는 것만)."""
    tmap = {t["코드"]: t["대사"] for t in tb_rows}
    grp = {t["코드"]: t["그룹"] for t in tb_rows}
    mv = {}
    for x in lines:
        j = judged[x["키"]]
        if j["마감"]:
            continue
        k = tmap.get(x["코드"], x["코드"])
        if grp.get(k) == "제조" and j["원가"]:
            continue
        mv[k] = mv.get(k, 0.0) + x["차변"] - x["대변"]
    out = {}
    re_open = re_close = re_mv = 0.0
    for t in tb_rows:
        g = t["그룹"]
        if g == "이익잉여금":
            re_open += t["기초"]
            re_close += t["기말"]
            continue
        if t["대사"] != t["코드"]:
            continue
        op = t["기초"] if g in ("재무상태표", "미분류") else 0.0
        d = t["기말"] - (op + mv.get(t["코드"], 0.0))
        if abs(d) > 0.5:
            out[str(t["코드"])] = round(d, 2)
    for t in tb_rows:
        if t["그룹"] == "이익잉여금":
            re_mv += mv.get(t["코드"], 0.0)
    pl_open = sum(t["기초"] for t in tb_rows if t["그룹"] == "손익")
    pl_mv = 0.0
    for x in lines:
        j = judged[x["키"]]
        k = tmap.get(x["코드"], x["코드"])
        if not j["마감"] and grp.get(k) in ("손익", "제조"):
            pl_mv += x["차변"] - x["대변"]
    if any(t["그룹"] == "이익잉여금" for t in tb_rows):
        d = re_close - (re_open + (pl_open if prior_add == "O" else 0) + re_mv + (pl_mv if cur_in == "O" else 0))
        if abs(d) > 0.5:
            out["이익잉여금"] = round(d, 2)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--journal", nargs="+", required=True, help="분개장 파일 · 폴더 · 와일드카드(여러 개면 합침)")
    ap.add_argument("--tb", default="")
    ap.add_argument("--tb-prior", default="")
    ap.add_argument("--tb-dir", default="")
    ap.add_argument("--tb-date", default="")
    ap.add_argument("--tb-prior-date", default="")
    ap.add_argument("--tb-json", default="")
    ap.add_argument("--company", required=True)
    ap.add_argument("--year-end", required=True, help="결산일 2026.12.31")
    ap.add_argument("--gaap", choices=["ifrs", "kgaap"], default="ifrs")
    ap.add_argument("--start", default="", help="대상 기간 시작(비우면 결산 연도 1.1)")
    ap.add_argument("--end", default="", help="기준일(비우면 분개장 마지막 날)")
    ap.add_argument("--pm", type=float, default=0)
    ap.add_argument("--author", default="")
    ap.add_argument("--wp-no", default="3600D")
    ap.add_argument("--list-max", type=int, default=100)
    ap.add_argument("--out", required=True)
    ap.add_argument("--overwrite", action="store_true")
    a = ap.parse_args()

    if os.path.exists(a.out) and not a.overwrite:
        raise KicpaError("이미 있습니다(사용자가 고쳤을 수 있음 - 확인 후 --overwrite 또는 다른 --out): %s" % a.out)
    ye = TBM.date_key(a.year_end)
    start = TBM.date_key(a.start) if a.start else dt.date(ye.year, 1, 1)
    end = TBM.date_key(a.end) if a.end else None
    lines, info = J.parse_many(a.journal, start=start, end=end, year=ye.year)
    if len(lines) > MAX_LINES:
        raise KicpaError("분개장 줄이 %s 줄로 엑셀 한 시트(1,048,576행)에 넣기 어렵습니다 - 기간을 나눠 만듭니다" % "{:,}".format(len(lines)))
    vouchers = J.vouchers_of(lines)
    tb, tbinfo = TBM.load(a.tb or None, a.tb_dir or None, a.tb_json or None, a.tb_date or None, a.tb_prior_date or None,
                          a.tb_prior or None)
    tb_rows, warns = tb_rows_of(tb, lines)
    prior_add, cur_in = re_toggles(tb_rows)
    info.update({"회사명": a.company, "결산일": W.fmt_date(ye), "작성자": a.author, "gaap": a.gaap, "pm_given": bool(a.pm),
                 "전기손익가산": prior_add, "당기손익포함": cur_in,
                 "tb": "%s %s" % (tbinfo["형식"], tbinfo["파일"]) + (" (%s / %s)" % (tbinfo["당기일"], tbinfo["전기일"]) if tbinfo.get("당기일") else "")})
    if a.pm:
        pm = int(a.pm)
    else:
        assets = sum(t["기말"] for t in tb_rows if t["구분"] == "자산")
        rev = -sum(t["기말"] for t in tb_rows if C.is_revenue(t["코드"], t["계정과목"]))
        months = (info["끝"].year - info["시작"].year) * 12 + info["끝"].month - info["시작"].month + 1
        pm = max(int(max(assets, rev * 12 / max(months, 1)) * 0.005 * 0.75) // 1000000 * 1000000, 1000000)
        warns.append("수행중요성을 주지 않아 임시값 %s(자산 · 연환산 매출 중 큰 값 × 0.5%% × 75%%)을 넣었습니다" % "{:,}".format(pm))
    holidays = json.load(open(os.path.join(HERE, "holidays.json"), encoding="utf-8"))
    has_wd = bool(info["열"].get("작성일"))
    P = R.default_params(lines, info, pm, holidays, has_wd)
    judged = R.judge(lines, vouchers, tb_rows, P)

    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    build(lines, vouchers, tb_rows, info, P, judged, a.out, a.wp_no, a.list_max)

    comp = completeness(lines, tb_rows, judged, prior_add, cur_in)
    summ = {"파일": a.out, "분개장": info["파일"], "완전성 차이": comp, "형식": info["형식"], "기간": [info["시작"].isoformat(), info["끝"].isoformat()],
            "줄 수": len(lines), "전표 수": len(vouchers), "차변": sum(x["차변"] for x in lines), "대변": sum(x["대변"] for x in lines),
            "수행중요성": pm, "임시 PM": not a.pm, "전기손익가산": prior_add, "당기손익포함": cur_in,
            "마감 전표": sum(j["마감"] for j in judged.values()), "원가대체 전표": sum(j["원가"] for j in judged.values()),
            "시나리오": {s: {"전표": sum(j[s] for j in judged.values()),
                          "줄": sum(v["e"] - v["s"] + 1 for v in vouchers if judged[v["키"]][s])} for s in R.SCN},
            "추출 전표(합집합)": sum(1 for j in judged.values() if any(j[s] for s in R.SCN)),
            "기간 밖 제외 줄": info["기간 밖 제외"], "경고": warns, "wp_no": a.wp_no}
    sp = os.path.splitext(a.out)[0] + "_summary.json"
    json.dump(summ, open(sp, "w", encoding="utf-8"), ensure_ascii=False, indent=1, default=str)
    print("저장: %s" % a.out)
    print("분개장 %s(%s) %s ~ %s  줄 %s · 전표 %s · 차변 %s · 대변 %s" % (
        info["파일"], info["형식"], info["시작"], info["끝"], "{:,}".format(len(lines)), "{:,}".format(len(vouchers)),
        "{:,.0f}".format(summ["차변"]), "{:,.0f}".format(summ["대변"])))
    print("수행중요성 %s%s · 마감 전표 %d · 원가대체 전표 %d · 이익잉여금 대사(전기 손익 더함 %s · 당기 손익 포함 %s)" % (
        "{:,}".format(pm), " (임시)" if not a.pm else "", summ["마감 전표"], summ["원가대체 전표"], prior_add, cur_in))
    for s in R.SCN:
        print("  %s %-14s 전표 %5d · 줄 %6d%s" % (s, R.SCN_INFO[s][1], summ["시나리오"][s]["전표"], summ["시나리오"][s]["줄"],
                                             "" if P["적용"][s] == "O" else "  (적용 X)"))
    print("  추출 전표(합집합) %d" % summ["추출 전표(합집합)"])
    for w in warns:
        print("경고: " + w)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        main()
    except KicpaError as ex:
        print("ERROR\t%s" % ex)
        sys.exit(2)
