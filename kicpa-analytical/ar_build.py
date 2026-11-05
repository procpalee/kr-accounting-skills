# -*- coding: utf-8 -*-
"""kicpa-analytical - 한공회 표준조서 2120A 위험평가 분석적절차(계획) · 8110A 전반적인 결론을 위한 분석적절차(종결) 조서.

  python -X utf8 ar_build.py prepare|build --stage 계획|기말 --gaap ifrs|kgaap --company <회사명> --year-end 2026.12.31
        (--tb <xlsx|csv> [--tb-prior <xlsx|csv>] | --tb-dir <xlaudit tb 폴더> --tb-date 2026.09.30 --tb-prior-date 2025.12.31
         | --tb-json <douzone_tb json>)
        [--journal <분개장 파일 · 폴더> ...] [--pm 819000000] [--ratio 0.1] [--annualize O|X] [--author 이름]
        [--notes notes.json] [--wp-no 2120A] --out <xlsx> [--overwrite]

prepare  분석 결과(analysis.json)와 AI 초안 작성용 근거(analysis.md - 유의적 차이 과목마다 TB 계정 · 분개장 상대계정 · 거래처)를
         --out 옆 '<조서>_work' 폴더에 쓴다. 이 대화의 Claude 가 analysis.md 를 읽고 notes.json(과목키 → 설명 문장)을 쓴다.
build    조서를 만든다(--notes 를 주면 '설명 또는 감사계획에 반영할 사항' 칸에 초안이 들어감 - 흰 칸, 사람이 고쳐 씀).

시트: 기초정보 · 2120A(또는 8110A) · 2120A-1 재무비율 · [2120A-2 월별 추세] · Data_TB · Data_과목 · [Data_분개장].
금액 · 백분율 · 증감 · 유의적 차이(잠정 O/△/X) · 비율은 모두 엑셀 수식이고, 계정 → 과목 매핑은 Data_TB '과목키' 칸(크림)에서 고친다.
"""
import argparse
import datetime as dt
import json
import os
import re
import sys
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
CORE = os.path.join(os.path.dirname(HERE), "kicpa-jet")
sys.path.insert(0, CORE)

from kicpa_core import KicpaError, coa as C, frameworks as F, journal as J, tb as TBM, wp as W  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402
from openpyxl.utils import get_column_letter as CL  # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402
from excel_theme import write_audit_header, write_section_bar, write_table, plain_punct  # noqa: E402

TH, ACC, PCT = W.TH, W.ACC, "percent_acct"
DATA_ROW = 6
GAAP = {"ifrs": "한국채택국제회계기준", "kgaap": "일반기업회계기준"}
STAGE = {
    "계획": {"wp": "2120A", "name": "위험평가 분석적절차", "std": "기준서 315",
             "purpose": "전체재무제표 수준 및 경영진주장 수준의 중요왜곡표시위험을 식별하고 평가하기 위하여 위험평가 분석적절차를 수행한다(한공회 2120).",
             "concl": "위험평가 분석적절차에 따른 결론"},
    "기말": {"wp": "8110A", "name": "전반적인 결론을 위한 분석적절차", "std": "기준서 520",
             "purpose": "감사의 종료에 근접하여 기업에 대하여 이해한 바와 재무제표가 일치하는지 전반적인 결론을 내리기 위하여 분석적절차를 수행한다(한공회 8110).",
             "concl": "전반적인 결론을 위한 분석적절차에 따른 결론"},
}
SEC_TITLE = {"BS": "재무상태표", "IS": "손익계산서", "MC": "제조원가명세서"}


def ratios(fw):
    """(구분, 비율, 산식 설명, 식{키}, 형식). 식은 당기_ · 전기_ 이름으로 바뀐다."""
    if fw == "ifrs":
        debt = "({단기차입금}+{장기차입금}+{유동리스}+{비유동리스})"
        intr = "{금융원가}"
        net = "({단기차입금}+{장기차입금}+{유동리스}+{비유동리스}-{현금}-{단기금융})"
    else:
        debt = "({단기차입금}+{유동성장기}+{사채}+{장기차입금})"
        intr = "{이자비용}"
        net = "({단기차입금}+{유동성장기}+{사채}+{장기차입금}-{현금})"
    out = [
        ("안정성", "유동비율", "유동자산 ÷ 유동부채", "{유동자산}/{유동부채}", PCT),
        ("안정성", "당좌비율", "(유동자산 - 재고자산) ÷ 유동부채", "({유동자산}-{재고})/{유동부채}", PCT),
        ("안정성", "부채비율", "부채총계 ÷ 자본총계", "{부채총계}/{자본총계}", PCT),
        ("안정성", "차입금의존도", "차입금" + ("(리스부채 포함)" if fw == "ifrs" else "") + " ÷ 자산총계", debt + "/{자산총계}", PCT),
        ("안정성", "순차입금", "차입금 - 현금성 자산", net, ACC),
        ("안정성", "이자보상배율", "영업이익 ÷ " + ("금융원가" if fw == "ifrs" else "이자비용"), "{영업이익}/" + intr, "0.00"),
        ("수익성", "매출총이익률", "매출총이익 ÷ 매출액", "{매출총이익}/{매출액}", PCT),
        ("수익성", "영업이익률", "영업이익 ÷ 매출액", "{영업이익}/{매출액}", PCT),
        ("수익성", "매출액순이익률", "당기순이익 ÷ 매출액", "{당기순이익}/{매출액}", PCT),
        ("수익성", "총자본이익률(ROA)", "당기순이익 ÷ 자산총계", "{당기순이익}/{자산총계}", PCT),
        ("수익성", "자기자본이익률(ROE)", "당기순이익 ÷ 자본총계", "{당기순이익}/{자본총계}", PCT),
        ("활동성", "총자본회전율", "매출액 ÷ 자산총계(회)", "{매출액}/{자산총계}", "0.00"),
        ("활동성", "자기자본회전율", "매출액 ÷ 자본총계(회)", "{매출액}/{자본총계}", "0.00"),
        ("활동성", "매출채권회전율", "매출액 ÷ 매출채권(회)", "{매출액}/{매출채권}", "0.00"),
        ("활동성", "매출채권 회수기일", "매출채권 × 365 ÷ 매출액(일)", "{매출채권}*365/{매출액}", "0.0"),
        ("활동성", "재고자산회전율", "매출원가 ÷ 재고자산(회)", "{매출원가}/{재고}", "0.00"),
        ("활동성", "매입채무 회전기간", "매입채무 × 365 ÷ 매출원가(일)", "{매입채무}*365/{매출원가}", "0.0"),
    ]
    return out


def load(a):
    ye = TBM.date_key(a.year_end)
    tb, tbinfo = TBM.load(a.tb or None, a.tb_dir or None, a.tb_json or None, a.tb_date or None, a.tb_prior_date or None,
                          a.tb_prior or None)
    cur = TBM.date_key(a.tb_date) if a.tb_date else ye
    prior = TBM.date_key(a.tb_prior_date) if a.tb_prior_date else dt.date(ye.year - 1, ye.month, ye.day)
    rows = []
    for t in tb:
        sec, key, why = F.map_row(a.gaap, t)
        rows.append(dict(t, 구역=sec or "", 과목키=key or "", 근거=why, 구분=C.kind_of(t["코드"], t["계정과목"], t.get("구분"))))
    fy_start = dt.date(ye.year - 1, ye.month, ye.day) + dt.timedelta(days=1)
    months = (cur.year - fy_start.year) * 12 + cur.month - fy_start.month + 1
    months = max(1, min(12, months))
    ann = a.annualize or ("O" if a.stage == "계획" and months < 12 and (prior.month, prior.day) == (ye.month, ye.day) else "X")
    bs = lambda f: sum(r[f] for r in rows if r["구역"] == "BS")
    pl = lambda f: sum(r[f] for r in rows if r["구역"] == "IS")      # 순이익 = 손익 계정만(제조원가는 매출원가에 이미 들어 있음)
    ni_add = "O" if abs(bs("기말")) > 1 and abs(bs("기말") + pl("기말")) <= 1 else "X"
    warns = []
    ni_add_p = "O" if abs(bs("기초")) > 1 and abs(bs("기초") + pl("기초")) <= 1 else "X"
    if ni_add_p != ni_add:
        warns.append("당기 · 전기 TB 의 손익 처리가 다릅니다(당기 순이익 가산 %s · 전기 %s) - 기초정보에서 확인" % (ni_add, ni_add_p))
    for r in rows:
        if r["근거"] == "기본값" and (r["기말"] or r["기초"]):
            warns.append("기본값 과목 배정: %s %s → %s (Data_TB 과목키 확인)" % (r["코드"], r["계정과목"], r["과목키"]))
        if not r["구역"]:
            warns.append("계정 성격 미분류: %s %s (TB 구분 열 필요)" % (r["코드"], r["계정과목"]))
    info = {"회사명": a.company, "결산일": W.fmt_date(ye), "기준일": cur, "비교일": prior, "경과월수": months, "연환산": ann,
            "순이익가산": ni_add, "작성자": a.author, "gaap": a.gaap, "stage": a.stage,
            "tb": "%s %s" % (tbinfo["형식"], tbinfo["파일"])}
    pm = int(a.pm) if a.pm else None
    if pm is None:
        assets = sum(r["기말"] for r in rows if r["구분"] == "자산")
        rev = -sum(r["기말"] for r in rows if r["과목키"] == "매출액")
        pm = max(int(max(assets, rev * (12 / months if ann == "O" else 1)) * 0.005 * 0.75) // 1000000 * 1000000, 1000000)
        warns.append("수행중요성을 주지 않아 임시값 %s 를 넣었습니다(2700 확정 후 기초정보에서 고침)" % "{:,}".format(pm))
    info["pm"], info["pm_given"] = pm, bool(a.pm)
    return rows, info, warns


def compute(rows, info, fw, ratio):
    """조서 수식과 같은 계산(파이썬) - analysis.json · 대조용."""
    fac = 12 / info["경과월수"] if info["연환산"] == "O" else 1
    val = {}
    for sec in ("BS", "IS", "MC"):
        for k, l, lv, t, f, rx in F.FRAMEWORKS[fw][sec]:
            if t != "leaf":
                continue
            sg = F.sign_of(fw, sec, k)
            for side, fld in (("당기", "기말"), ("전기", "기초")):
                v = sg * sum(r[fld] for r in rows if r["구역"] == sec and r["과목키"] == k)
                if sec in ("IS", "MC") and side == "당기":
                    v *= fac
                val[(sec, k, side)] = v
    for side, fld in (("당기", "기말"), ("전기", "기초")):
        ni = -sum(r[fld] for r in rows if r["구역"] == "IS")
        if info["순이익가산"] == "O":
            rek = "이익잉여금" if fw == "ifrs" else "미처분"
            val[("BS", rek, side)] += ni
    for sec in ("BS", "IS", "MC"):
        for side in ("당기", "전기"):
            todo = [x for x in F.FRAMEWORKS[fw][sec] if x[3] in ("sum", "calc", "ratio")]
            while todo:                                  # 하위 줄이 아래에 있어도 풀리도록 되풀이
                left = []
                for k, l, lv, t, f, rx in todo:
                    deps = f if t == "sum" else re.findall(r"\{(\w+)\}", f)
                    if any((sec, c, side) not in val for c in deps):
                        left.append((k, l, lv, t, f, rx))
                        continue
                    if t == "sum":
                        val[(sec, k, side)] = sum(val[(sec, c, side)] for c in f)
                    else:
                        expr = re.sub(r"\{(\w+)\}", lambda m: repr(val[(sec, m.group(1), side)]), f)
                        try:
                            val[(sec, k, side)] = eval(expr)
                        except (ZeroDivisionError, TypeError):
                            val[(sec, k, side)] = None
                if len(left) == len(todo):
                    raise KicpaError("과목 골격 식을 풀지 못했습니다: %s" % [x[0] for x in left])
                todo = left
    lines = []
    for sec in ("BS", "IS", "MC"):
        for k, l, lv, t, f, rx in F.FRAMEWORKS[fw][sec]:
            a, b = val[(sec, k, "당기")], val[(sec, k, "전기")]
            if t == "ratio" or a is None or b is None:
                continue
            d = a - b
            rt = d / abs(b) if b else None
            flag = ""
            if t in ("leaf", "sum") and (a or b):
                big = abs(d) >= info["pm"]
                bigr = rt is not None and abs(rt) >= ratio
                flag = "O" if big and (b == 0 or bigr) else ("△" if big or bigr else "X")
            lines.append({"구역": sec, "키": k, "과목": l, "종류": t, "당기": a, "전기": b, "증감": d, "증감률": rt, "잠정": flag})
    return lines, val


# ──────────────────────────────────────────────────────────────
# prepare - AI 초안 근거
# ──────────────────────────────────────────────────────────────
def evidence(rows, lines, journal, fw):
    if not journal:
        return {}
    L = journal["lines"]
    vou = defaultdict(list)
    for x in L:
        vou[x["키"]].append(x)
    tbc = {r["코드"] for r in rows}
    par = {}
    for x in L:
        c = x["코드"]
        if c not in par:
            p = c
            if c not in tbc and isinstance(c, int):
                s = str(c)
                for k in range(1, 5):
                    q = int(s[:-k] + "0" * k)
                    if q in tbc:
                        p = q
                        break
            par[c] = p
    out = {}
    for ln in lines:
        if ln["잠정"] not in ("O", "△") or ln["종류"] != "leaf":
            continue
        codes = {r["코드"] for r in rows if (r["구역"], r["과목키"]) == (ln["구역"], ln["키"])}
        mine = [x for x in L if par.get(x["코드"]) in codes]
        if not mine:
            continue
        keys = {x["키"] for x in mine}
        cp, cust, big = Counter(), Counter(), []
        for k in keys:
            vs = vou[k]
            amt = sum(x["차변"] for x in vs)
            big.append((amt, k, next((x["적요"] for x in vs if x["적요"]), "")))
            for x in vs:
                if par.get(x["코드"]) not in codes:
                    cp[(x["코드"], x["계정과목"])] += x["차변"] + x["대변"]
            for x in vs:
                if par.get(x["코드"]) in codes and x["거래처명"]:
                    cust[x["거래처명"]] += x["차변"] + x["대변"]
        out[(ln["구역"], ln["키"])] = {
            "분개 줄": len(mine), "전표": len(keys),
            "차변": sum(x["차변"] for x in mine), "대변": sum(x["대변"] for x in mine),
            "상대계정": [("%s %s" % c, v) for c, v in cp.most_common(5)],
            "거래처": cust.most_common(5),
            "큰 전표": sorted(big, reverse=True)[:3]}
    return out


def write_prepare(rows, info, lines, ev, workdir, fw):
    os.makedirs(workdir, exist_ok=True)
    json.dump({"info": {k: (v.isoformat() if isinstance(v, dt.date) else v) for k, v in info.items()}, "lines": lines},
              open(os.path.join(workdir, "analysis.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    f = lambda v: "{:,.0f}".format(v) if isinstance(v, (int, float)) else str(v)
    md = ["# %s %s - AI 초안 근거" % (STAGE[info["stage"]]["wp"], info["회사명"]), "",
          "- 회계기준 %s · 당기 %s · 비교 %s · 경과 %d개월 · 손익 연환산 %s · 판단기준 금액 %s" % (
              GAAP[fw], info["기준일"], info["비교일"], info["경과월수"], info["연환산"], f(info["pm"])),
          "", "## 쓰는 법", "",
          "유의적 차이(잠정 O · △) 과목마다 '설명 또는 감사계획에 반영할 사항' 초안을 notes.json 에 쓴다.",
          "형식: {\"설명\": {\"BS:매출채권\": \"문장\", \"IS:매출액\": \"문장\"}}. 개조식 1~3줄, 끝에 마침표 없음.",
          "자료로 확인한 사실(증감 금액 · 상대계정 · 거래처 · 큰 전표)만 쓰고, 원인을 모르면 '회사 질문 필요 - …'로 쓴다. 결론 · 최종 O/X 는 사람이 정한다.",
          "문체는 stop-slop-ko 규칙(번역투 · 상투구 · 줄표 없이)."]
    for ln in lines:
        if ln["잠정"] not in ("O", "△"):
            continue
        md += ["", "## %s:%s %s (잠정 %s)" % (ln["구역"], ln["키"], ln["과목"], ln["잠정"]),
               "- 당기 %s · 전기 %s · 증감 %s · 증감률 %s" % (f(ln["당기"]), f(ln["전기"]), f(ln["증감"]),
                                                     "%.1f%%" % (ln["증감률"] * 100) if ln["증감률"] is not None else "-")]
        if ln["종류"] == "leaf":
            accts = [r for r in rows if (r["구역"], r["과목키"]) == (ln["구역"], ln["키"])]
            accts.sort(key=lambda r: -abs(r["기말"] - r["기초"]))
            for r in accts[:8]:
                md.append("  - TB %s %s: 당기 %s · 전기 %s" % (r["코드"], r["계정과목"], f(r["기말"]), f(r["기초"])))
        e = ev.get((ln["구역"], ln["키"]))
        if e:
            md.append("- 분개장: 줄 %s · 전표 %s · 차변 %s · 대변 %s" % (f(e["분개 줄"]), f(e["전표"]), f(e["차변"]), f(e["대변"])))
            md.append("  - 상대계정: " + ", ".join("%s %s" % (c, f(v)) for c, v in e["상대계정"]))
            if e["거래처"]:
                md.append("  - 거래처: " + ", ".join("%s %s" % (c, f(v)) for c, v in e["거래처"]))
            md.append("  - 큰 전표: " + ", ".join("%s %s(%s)" % (k, f(v), m[:30]) for v, k, m in e["큰 전표"]))
    open(os.path.join(workdir, "analysis.md"), "w", encoding="utf-8").write("\n".join(md) + "\n")


# ──────────────────────────────────────────────────────────────
# build
# ──────────────────────────────────────────────────────────────
def build(rows, info, fw, ratio, notes, journal, out_path, wp_no):
    st = STAGE[info["stage"]]
    wb = Workbook()
    wb.remove(wb.active)
    sname = {"g": "기초정보", "m": wp_no, "r": "%s-1 재무비율" % wp_no, "t": "%s-2 월별 추세" % wp_no,
             "dt": "Data_TB", "dk": "Data_과목", "dj": "Data_분개장"}
    keys = ["g", "m", "r"] + (["t"] if journal else []) + ["dt", "dk"] + (["dj"] if journal else [])
    ws = {k: wb.create_sheet(sname[k]) for k in keys}
    nm = lambda n, sh, ref: W.nm(wb, n, sh, ref)
    meta = lambda no: {"회사명": info["회사명"], "조서번호": "%s-%s" % (wp_no, no) if no else wp_no, "결산일": info["결산일"],
                       "조서명": st["name"], "작성자": info.get("작성자", "")}
    is8110 = info["stage"] == "기말"

    # ── Data_과목 ──
    sk = ws["dk"]
    write_audit_header(sk, meta(None), theme=TH, tab="calc")
    klist = []
    for sec in ("BS", "IS", "MC"):
        for k, l, lv, t, f, rx in F.FRAMEWORKS[fw][sec]:
            if t == "leaf":
                klist.append([sec, k, l, F.sign_of(fw, sec, k)])
    write_table(sk, ["구역", "과목키", "표시명", "부호"], klist, theme=TH, header_row=5, number_cols={3: "0"})
    nm("과목_키", sname["dk"], "$C$6:$C$%d" % (5 + len(klist)))
    for c, w in (("B", 8), ("C", 16), ("D", 38), ("E", 8)):
        sk.column_dimensions[c].width = w

    # ── Data_TB ──
    sd = ws["dt"]
    write_audit_header(sd, meta(None), theme=TH, tab="pbc")
    drow = [[r["코드"], r["계정과목"], r["구분"], r.get("표시과목"), r["구역"], r["과목키"], r["근거"], r["기말"], r["기초"]] for r in rows]
    write_table(sd, ["계정코드", "계정과목", "구분", "표시과목", "구역", "과목키", "매핑 근거", "당기(TB)", "전기(TB)"], drow, theme=TH,
                header_row=5, number_cols={0: "0", 7: ACC, 8: ACC}, input_cells=["F6:G%d" % (5 + len(drow))])
    tl = 5 + len(drow)
    for n, c in (("TB_구역", "F"), ("TB_키", "G"), ("TB_당기", "I"), ("TB_전기", "J"), ("TB_코드", "B")):
        nm(n, sname["dt"], "$%s$6:$%s$%d" % (c, c, tl))
    dv = DataValidation(type="list", formula1="=과목_키", allow_blank=True)
    sd.add_data_validation(dv)
    dv.add("G6:G%d" % tl)
    for c, w in (("B", 11), ("C", 34), ("D", 10), ("E", 34), ("F", 8), ("G", 16), ("H", 12)):
        sd.column_dimensions[c].width = w

    # ── 기초정보 ──
    sg = ws["g"]
    hdr = write_audit_header(sg, meta(None), sheet_title="기초정보", theme=TH, tab="input")
    r = write_section_bar(sg, hdr, "1. 조서 목적", theme=TH, last_col=6)
    W.put(sg, r - 1, 2, st["purpose"])
    r += 1
    r = write_section_bar(sg, r, "2. 기본 정보", theme=TH, last_col=6)
    base = [("회사명", info["회사명"], None), ("결산일", info["결산일"], None), ("회계기준", GAAP[fw], None),
            ("단계", "%s(%s)" % (info["stage"], st["std"]), None),
            ("당기 기준일", info["기준일"], "당기기준일"), ("비교 기준일", info["비교일"], "전기기준일"),
            ("경과 월수", info["경과월수"], "경과월수"), ("손익 연환산(O/X)", info["연환산"], "연환산"),
            ("연환산 계수", '=IF(연환산="O",12/경과월수,1)', "연환산계수"),
            ("이익잉여금에 당기순이익 더함(O/X)", info["순이익가산"], "순이익가산"),
            ("TB", info["tb"], None), ("작성자", info.get("작성자") or None, None)]
    write_table(sg, ["항목", "값"], [[a, b] for a, b, c in base], theme=TH, header_row=r, number_cols={},
                input_cells=["C%d" % (r + 1 + i) for i, x in enumerate(base) if x[0] in ("손익 연환산(O/X)", "이익잉여금에 당기순이익 더함(O/X)", "작성자")])
    for i, (a, b, n) in enumerate(base):
        c = sg["C%d" % (r + 1 + i)]
        if n:
            nm(n, sname["g"], "$C$%d" % (r + 1 + i))
        if isinstance(b, dt.date):
            c.number_format = "yyyy-mm-dd"
        elif a == "연환산 계수":
            c.number_format = "0.0000"
        c.alignment = Alignment(horizontal="right" if isinstance(b, (int, float, dt.date)) or a == "연환산 계수" else "left",
                                vertical="center")
    r = r + len(base) + 2
    r = write_section_bar(sg, r, "3. 판단 기준", theme=TH, last_col=6)
    jb = [("수행중요성", info["pm"], "2700 중요성 조서" + ("" if info["pm_given"] else " - 임시값, 확정 후 고칠 것"), "수행중요성"),
          ("계정의 유의성 판단기준(금액)", "=수행중요성", "기본 = 수행중요성", "판단금액"),
          ("계정의 유의성 판단기준(비율)", ratio, "증감률 기준", "판단비율")]
    write_table(sg, ["항목", "값", "설명"], [list(x[:3]) for x in jb], theme=TH, header_row=r, number_cols={1: ACC},
                input_cells=["C%d:C%d" % (r + 1, r + 3)])
    sg["C%d" % (r + 3)].number_format = "0%"
    for i, x in enumerate(jb):
        nm(x[3], sname["g"], "$C$%d" % (r + 1 + i))
    sg.column_dimensions["B"].width = 40
    sg.column_dimensions["C"].width = 30
    sg.column_dimensions["D"].width = 30

    # ── 본 시트(2120A · 8110A) ──
    sm = ws["m"]
    hdr = write_audit_header(sm, meta(None), sheet_title=st["name"], theme=TH, tab="output")
    head = ["과목", "당기 금액", "당기 백분율", "전기 금액", "전기 백분율", "증감 금액", "증감률", "유의적 차이(잠정)", "유의적 차이(최종)",
            "설명 또는 감사계획에 반영할 사항", ""] + (["계획단계 식별 사항 해소"] if is8110 else []) + ["관련 조서번호"]
    last = 1 + len(head)
    r = write_section_bar(sm, hdr, "1. 판단 기준", theme=TH, last_col=last)
    write_table(sm, ["항목", "값"], [["계정의 유의성 판단기준(금액)", "=판단금액"], ["계정의 유의성 판단기준(비율)", "=판단비율"],
                                   ["당기 기준일", "=당기기준일"], ["비교 기준일", "=전기기준일"],
                                   ["손익 연환산", '=IF(연환산="O","O(× "&TEXT(연환산계수,"0.00")&")","X")']],
                theme=TH, header_row=r, number_cols={1: ACC})
    sm["C%d" % (r + 2)].number_format = "0%"
    for i in (3, 4):
        sm["C%d" % (r + i)].number_format = "yyyy-mm-dd"
    for i in (3, 4):
        sm["C%d" % (r + i)].alignment = Alignment(horizontal="right", vertical="center")
    sm["C%d" % (r + 5)].alignment = Alignment(horizontal="left", vertical="center")
    r += 7
    rowof = {}
    note_cells = []
    sec_no = 2
    for sec in ("BS", "IS", "MC"):
        lines = F.FRAMEWORKS[fw][sec]
        if sec == "MC" and not any(x["구역"] == "MC" for x in rows):
            continue
        r = write_section_bar(sm, r, "%d. %s" % (sec_no, SEC_TITLE[sec]), theme=TH, last_col=last)
        sec_no += 1
        hr = r
        fr = hr + 1
        for i, x in enumerate(lines):
            rowof[(sec, x[0])] = fr + i
        base_key = {"BS": None, "IS": "매출액", "MC": "총제조비용"}[sec]
        body, totals, subs = [], [], []
        for i, (k, l, lv, t, f, rx) in enumerate(lines):
            rr = fr + i
            ref = lambda kk, col: "%s%d" % (col, rowof[(sec, kk)])
            if t == "leaf":
                sg_ = F.sign_of(fw, sec, k)
                cur = "%sSUMIFS(TB_당기,TB_구역,\"%s\",TB_키,\"%s\")" % ("-" if sg_ < 0 else "", sec, k)
                pri = cur.replace("TB_당기", "TB_전기")
                if sec in ("IS", "MC"):
                    cur = "(%s)*연환산계수" % cur
                if sec == "BS" and k == ("이익잉여금" if fw == "ifrs" else "미처분"):
                    ni = '-SUMIFS(TB_당기,TB_구역,"IS")'
                    cur += '+IF(순이익가산="O",%s,0)' % ni
                    pri += '+IF(순이익가산="O",%s,0)' % ni.replace("TB_당기", "TB_전기")
                cf, pf = "=" + cur, "=" + pri
            elif t == "sum":
                cf = "=" + "+".join(ref(c, "C") for c in f)
                pf = "=" + "+".join(ref(c, "E") for c in f)
            else:
                cf = "=" + re.sub(r"\{(\w+)\}", lambda m: ref(m.group(1), "C"), f)
                pf = "=" + re.sub(r"\{(\w+)\}", lambda m: ref(m.group(1), "E"), f)
                if t == "ratio":
                    cf = "=IFERROR(%s,\"\")" % cf[1:]
                    pf = "=IFERROR(%s,\"\")" % pf[1:]
            if sec == "BS":
                side = F.side_of(fw, k)
                den = "자산총계" if side == "자산" else "부채와자본총계"
            else:
                den = base_key
            pct = lambda col: '=IFERROR(%s%d/%s,"")' % (col, rr, ref(den, col))
            ratio_row = t == "ratio"
            g = "" if ratio_row else "=C%d-E%d" % (rr, rr)
            h = "" if ratio_row else '=IF(E%d=0,"",G%d/ABS(E%d))' % (rr, rr, rr)
            judge = ""
            if t in ("leaf", "sum"):
                judge = ('=IF(AND(C{r}=0,E{r}=0),"",IF(AND(ABS(G{r})>=판단금액,OR(E{r}=0,ABS(N(H{r}))>=판단비율)),"O",'
                         'IF(OR(ABS(G{r})>=판단금액,AND(E{r}<>0,ABS(N(H{r}))>=판단비율)),"△","X")))').format(r=rr)
            row = [l, cf, "" if ratio_row else pct("C"), pf, "" if ratio_row else pct("E"), g, h, judge, None,
                   notes.get("%s:%s" % (sec, k)), None] + ([None] if is8110 else []) + [None]
            body.append([v if v != "" else None for v in row])
            if t == "calc" and lv == 0:
                totals.append(i + 1)
            elif t == "sum" or (t in ("calc",) and lv > 0):
                subs.append(i + 1)
        write_table(sm, head, body, theme=TH, header_row=hr,
                    number_cols={1: ACC, 2: PCT, 3: ACC, 4: PCT, 5: ACC, 6: PCT},
                    input_cells=["J%d:J%d" % (fr, fr + len(lines) - 1)] + (["M%d:M%d" % (fr, fr + len(lines) - 1)] if is8110 else []),
                    subtotal_rows=subs, total_rows=totals)
        W.merge_text_col(sm, hr, fr, fr + len(lines) - 1, 11, 12)          # 설명 칸 = K:L 병합(긴 글은 행높이 도구로)
        for i, (k, l, lv, t, f, rx) in enumerate(lines):
            rr = fr + i
            if t == "ratio":
                for col in "CE":
                    sm["%s%d" % (col, rr)].number_format = "0.00%"
            sm["B%d" % rr].alignment = Alignment(horizontal="left", vertical="center", indent=lv)
            for col in "IJ":
                sm["%s%d" % (col, rr)].alignment = Alignment(horizontal="left", vertical="center")
            for col in ("K", "M", "N") if is8110 else ("K", "M"):
                sm["%s%d" % (col, rr)].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            if notes.get("%s:%s" % (sec, k)):
                note_cells.append("K%d" % rr)
            nm("당기_%s_%s" % (sec, k), sname["m"], "$C$%d" % rr)
            nm("전기_%s_%s" % (sec, k), sname["m"], "$E$%d" % rr)
        r = fr + len(lines) + 1
        W.put(sm, r - 1, 2, "* 백분율은 " + ("자산총계(부채·자본은 부채와자본총계) 대비" if sec == "BS" else
                                         ("매출액 대비" if sec == "IS" else "당기총제조비용 대비")) + " 비율, 증감 = 당기 - 전기")
        r += 1
    # 검증
    r = write_section_bar(sm, r, "%d. 검증" % sec_no, theme=TH, last_col=last)
    sec_no += 1
    ta, tl_ = rowof[("BS", "자산총계")], rowof[("BS", "부채와자본총계")]
    ni_r = rowof[("IS", "당기순이익")]
    chk = [["자산총계 - 부채와자본총계", "=C%d-C%d" % (ta, tl_), "=E%d-E%d" % (ta, tl_), "0 이어야 함"],
           ["당기순이익 - TB 손익 합계", '=C%d-(-SUMIFS(TB_당기,TB_구역,"IS"))*연환산계수' % ni_r,
            '=E%d-(-SUMIFS(TB_전기,TB_구역,"IS"))' % ni_r, "0 이어야 함(제조원가 계정은 매출원가에 들어 있어 빼고 셈)"],
           ["과목키가 빈 TB 금액", '=SUMIFS(TB_당기,TB_키,"")', '=SUMIFS(TB_전기,TB_키,"")', "0 이어야 함(Data_TB 과목키)"],
           ["과목키가 목록에 없는 TB 줄 수", '=SUMPRODUCT((TB_키<>"")*(COUNTIF(과목_키,TB_키)=0))', None, "0 이어야 함"]]
    write_table(sm, ["항목", "당기", "전기", "기준", "", "", ""], [x + [None, None, None] for x in chk], theme=TH, header_row=r,
                number_cols={1: ACC, 2: ACC})
    W.merge_text_col(sm, r, r + 1, r + len(chk), 5, 8)
    chk_row = r + 1
    r = r + len(chk) + 2
    r = write_section_bar(sm, r, "%d. 결론" % sec_no, theme=TH, last_col=last)
    write_table(sm, [st["concl"]] + [""] * (len(head) - 1), [[None] * len(head)], theme=TH, header_row=r,
                input_cells=["B%d" % (r + 1)])
    W.merge_text_col(sm, r, r + 1, r + 1, 2, last)
    widths = {"B": 40, "C": 18, "D": 11, "E": 18, "F": 11, "G": 18, "H": 10, "I": 10, "J": 10, "K": 40, "L": 30}
    if is8110:
        widths.update({"M": 30, "N": 12})
    else:
        widths["M"] = 12
    for c, w in widths.items():
        sm.column_dimensions[c].width = w

    # ── 재무비율 ──
    sr = ws["r"]
    hdr = write_audit_header(sr, meta("1"), theme=TH, tab="calc")
    r = write_section_bar(sr, hdr, "1. 주요 재무비율", theme=TH, last_col=8)
    keyname = {}
    for sec in ("BS", "IS"):
        for x in F.FRAMEWORKS[fw][sec]:
            keyname.setdefault(x[0], sec)
    body = []
    for g, n, how, expr, fmt in ratios(fw):
        rr = r + 1 + len(body)
        cur = re.sub(r"\{(\w+)\}", lambda m: "당기_%s_%s" % (keyname[m.group(1)], m.group(1)), expr)
        pri = re.sub(r"\{(\w+)\}", lambda m: "전기_%s_%s" % (keyname[m.group(1)], m.group(1)), expr)
        body.append([g, n, how, '=IFERROR(%s,"")' % cur, '=IFERROR(%s,"")' % pri, '=IFERROR(E%d-F%d,"")' % (rr, rr), None])
    rr = r + 1 + len(body)
    body.append(["성장성", "매출액증가율", "당기 매출액 ÷ 전기 매출액 - 1", '=IFERROR(당기_IS_매출액/전기_IS_매출액-1,"")', None, None, None])
    body.append(["성장성", "총자산증가율", "당기 자산총계 ÷ 전기 자산총계 - 1", '=IFERROR(당기_BS_자산총계/전기_BS_자산총계-1,"")', None, None, None])
    write_table(sr, ["구분", "비율", "산식", "당기", "전기", "증감", "해설"], body, theme=TH, header_row=r,
                input_cells=["H%d:H%d" % (r + 1, r + len(body))])
    allr = ratios(fw) + [(None, None, None, None, PCT), (None, None, None, None, PCT)]
    for i, x in enumerate(allr):
        for col in "EFG":
            c = sr["%s%d" % (col, r + 1 + i)]
            c.number_format = W_FMT.get(x[4], x[4])
            c.alignment = Alignment(horizontal="right", vertical="center")
        sr["H%d" % (r + 1 + i)].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    r = r + len(body) + 1
    W.put(sr, r, 2, "* 손익 항목은 %s 금액, 재무상태표 항목은 기준일 잔액(평균 잔액 아님)" % ("연환산" if info["연환산"] == "O" else "기간"))
    for c, w in (("B", 10), ("C", 24), ("D", 42), ("E", 16), ("F", 16), ("G", 16), ("H", 44)):
        sr.column_dimensions[c].width = w

    # ── 월별 추세 · Data_분개장(분개장이 있을 때) ──
    if journal:
        L = journal["lines"]
        tbc = {r_["코드"]: r_ for r_ in rows}
        sj = ws["dj"]
        write_audit_header(sj, meta(None), theme=TH, tab="pbc")
        keep_keys = {"매출액", "매출원가"}
        pick = []
        for x in L:
            c = x["코드"]
            p = c
            if c not in tbc and isinstance(c, int):
                s = str(c)
                for kk in range(1, 5):
                    q = int(s[:-kk] + "0" * kk)
                    if q in tbc:
                        p = q
                        break
            if p in tbc and tbc[p]["구역"] == "IS" and tbc[p]["과목키"] in keep_keys:
                pick.append((x, p))
        jrows = []
        for i, (x, p) in enumerate(pick):
            rr = DATA_ROW + i
            jrows.append([x["일자"], x["코드"], x["계정과목"], p, '=IFERROR(INDEX(TB_키,MATCH(E%d,TB_코드,0)),"")' % rr, x["차변"], x["대변"]])
        W.write_data_fast(sj, ["일자", "계정코드", "계정과목", "TB 계정", "과목키", "차변", "대변"], jrows, header_row=5,
                          number_cols={0: "date", 1: "0", 3: "0", 5: ACC, 6: ACC}, text_cols=(2, 4))
        jl = DATA_ROW + len(jrows) - 1
        for n, c in (("분개_일자", "B"), ("분개_키", "F"), ("분개_차변", "G"), ("분개_대변", "H")):
            nm(n, sname["dj"], "$%s$%d:$%s$%d" % (c, DATA_ROW, c, max(jl, DATA_ROW)))
        stt = ws["t"]
        hdr = write_audit_header(stt, meta("2"), theme=TH, tab="calc")
        r = write_section_bar(stt, hdr, "1. 월별 매출·원가", theme=TH, last_col=8)
        months = []
        d0 = dt.date(journal["시작"].year, journal["시작"].month, 1)
        while d0 <= journal["끝"]:
            months.append(d0)
            d0 = dt.date(d0.year + (d0.month == 12), d0.month % 12 + 1, 1)
        body = []
        for m in months:
            rr = r + 1 + len(body)
            rng = '분개_일자,">="&B{r},분개_일자,"<"&EDATE(B{r},1)'.format(r=rr)
            body.append([m, '=SUMIFS(분개_대변,분개_키,"매출액",%s)-SUMIFS(분개_차변,분개_키,"매출액",%s)' % (rng, rng),
                         '=SUMIFS(분개_차변,분개_키,"매출원가",%s)-SUMIFS(분개_대변,분개_키,"매출원가",%s)' % (rng, rng),
                         "=C%d-D%d" % (rr, rr), '=IFERROR(D%d/C%d,"")' % (rr, rr), None])
        n0 = r + 1
        body.append(["합계", "=SUM(C%d:C%d)" % (n0, n0 + len(months) - 1), "=SUM(D%d:D%d)" % (n0, n0 + len(months) - 1),
                     "=C%d-D%d" % (n0 + len(months), n0 + len(months)), '=IFERROR(D%d/C%d,"")' % (n0 + len(months), n0 + len(months)), None])
        write_table(stt, ["월", "매출액", "매출원가", "매출총이익", "원가율", "해설"], body, theme=TH, header_row=r,
                    number_cols={0: "yyyy-mm", 1: ACC, 2: ACC, 3: ACC, 4: PCT}, total_rows=[len(body)],
                    input_cells=["G%d:G%d" % (n0, n0 + len(months) - 1)])
        for i in range(len(body)):
            stt["G%d" % (n0 + i)].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            for col in "CDEF":                       # SUMIFS 조건의 & 때문에 글로 판정되지 않게
                stt["%s%d" % (col, n0 + i)].alignment = Alignment(horizontal="right", vertical="center")
        r = r + len(body) + 1
        W.put(stt, r, 2, "* 분개장(%s ~ %s)의 매출액·매출원가 과목 계정 분개를 월별로 모은 금액(마감 대체 전표 포함 여부 확인)"
              % (journal["시작"].isoformat(), journal["끝"].isoformat()))
        W.put(stt, r + 1, 2, "* 과목키는 Data_TB 매핑을 따른다 - 2120 2.(1) 매출액, 5.(4) 원가율 분석")
        for c, w in (("B", 10), ("C", 18), ("D", 18), ("E", 18), ("F", 10), ("G", 40)):
            stt.column_dimensions[c].width = w

    W.fit_headers(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return {"검증 행": chk_row, "rowof": {"%s:%s" % k: v for k, v in rowof.items()}, "note_cells": note_cells}


W_FMT = {ACC: "#,###,##0;[Red](#,###,##0);-", PCT: '0.00%_);[Red](0.00%);"-"_);@_)'}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["prepare", "build"])
    ap.add_argument("--stage", choices=["계획", "기말"], required=True)
    ap.add_argument("--gaap", choices=["ifrs", "kgaap"], required=True)
    ap.add_argument("--company", required=True)
    ap.add_argument("--year-end", required=True)
    ap.add_argument("--tb", default="")
    ap.add_argument("--tb-prior", default="")
    ap.add_argument("--tb-dir", default="")
    ap.add_argument("--tb-date", default="")
    ap.add_argument("--tb-prior-date", default="")
    ap.add_argument("--tb-json", default="")
    ap.add_argument("--journal", nargs="*", default=[])
    ap.add_argument("--pm", type=float, default=0)
    ap.add_argument("--ratio", type=float, default=0.1)
    ap.add_argument("--annualize", choices=["O", "X"], default=None)
    ap.add_argument("--author", default="")
    ap.add_argument("--notes", default="")
    ap.add_argument("--wp-no", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--no-fit", action="store_true", help="병합 칸 행높이 맞춤(엑셀 COM)을 건너뜀")
    a = ap.parse_args()
    wp_no = a.wp_no or STAGE[a.stage]["wp"]

    rows, info, warns = load(a)
    lines, val = compute(rows, info, a.gaap, a.ratio)
    journal = None
    if a.journal:
        ye = TBM.date_key(a.year_end)
        jl, ji = J.parse_many(a.journal, start=dt.date(ye.year - 1, ye.month, ye.day) + dt.timedelta(days=1), end=info["기준일"],
                              year=ye.year)
        journal = {"lines": jl, "시작": ji["시작"], "끝": ji["끝"], "파일": ji["파일"]}
    work = os.path.splitext(a.out)[0] + "_work"
    if a.cmd == "prepare":
        ev = evidence(rows, lines, journal, a.gaap)
        write_prepare(rows, info, lines, ev, work, a.gaap)
        sig = [x for x in lines if x["잠정"] in ("O", "△")]
        print("근거: %s" % os.path.join(work, "analysis.md"))
        print("유의적 차이 과목 %d개(O %d · △ %d) - analysis.md 를 읽고 notes.json 을 쓴 뒤 build --notes" % (
            len(sig), sum(x["잠정"] == "O" for x in sig), sum(x["잠정"] == "△" for x in sig)))
    else:
        if os.path.exists(a.out) and not a.overwrite:
            raise KicpaError("이미 있습니다(사용자가 고쳤을 수 있음 - 확인 후 --overwrite 또는 다른 --out): %s" % a.out)
        notes = {}
        if a.notes:
            notes = {k: plain_punct(v).replace(" · ", "·") for k, v in json.load(open(a.notes, encoding="utf-8")).get("설명", {}).items()}
        os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
        res = build(rows, info, a.gaap, a.ratio, notes, journal, a.out, wp_no)
        if not a.no_fit:
            print(W.fit_rows(a.out))
        os.makedirs(work, exist_ok=True)
        json.dump({"info": {k: (v.isoformat() if isinstance(v, dt.date) else v) for k, v in info.items()}, "lines": lines,
                   "rowof": res["rowof"], "검증 행": res["검증 행"]},
                  open(os.path.join(work, "summary.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        print("저장: %s" % a.out)
        print("AI 초안 칸 %d개 · 유의적 차이(잠정 O) %d · △ %d" % (len(res["note_cells"]), sum(x["잠정"] == "O" for x in lines),
                                                         sum(x["잠정"] == "△" for x in lines)))
    print("회계기준 %s · %s %s vs %s · 경과 %d개월 · 연환산 %s · 순이익 가산 %s · 판단기준 %s" % (
        GAAP[a.gaap], STAGE[a.stage]["wp"], info["기준일"], info["비교일"], info["경과월수"], info["연환산"], info["순이익가산"],
        "{:,}".format(info["pm"])))
    for x in lines:
        if x["키"] in ("자산총계", "부채와자본총계", "매출액", "영업이익", "당기순이익"):
            print("  %-8s 당기 %18s · 전기 %18s" % (x["과목"], "{:,.0f}".format(x["당기"]), "{:,.0f}".format(x["전기"])))
    for w in warns:
        print("경고: " + w)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        main()
    except KicpaError as ex:
        print("ERROR\t%s" % ex)
        sys.exit(2)
