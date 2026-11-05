# -*- coding: utf-8 -*-
"""kicpa-sap - 한공회 표준조서 3300A 실증적 분석절차(기준서 520) 조서를 입력 JSON 으로 만든다.

  python -X utf8 sap_build.py <input.json> [--out <xlsx>] [--overwrite]

기대치 모델: 기간(월 등) × 항목(유종 · 호실 · 차입금 등)마다  기대치 = 드라이버 × 파라미터 ÷ 나눔 계수
  드라이버 = 원자료 열 × 환산계수 × (비중 - 원자료가 합산 수량이면 비중 표로 나눔)
  파라미터 = 외부 자료(시장가격 · 공시 요율 · 계약 단가)를 우선
시트: 3300A(한공회 요약표 - 대상마다 한 줄) · Ref.n <대상>(기준서 520 9단계) · Ref.n-A 드라이버 · Ref.n-B 파라미터.
기대치 · 차이 · 판정은 모두 엑셀 수식이고 원자료 · 파라미터 · 판단 문장만 값이다. 입력 스키마는 examples/ 의 예시 JSON.
산출 옆 <조서>_summary.json = 파이썬으로 계산한 기대치 · 차이(sap_check.py 가 엑셀 재계산 값과 대조).
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "kicpa-jet"))

from kicpa_core import KicpaError, wp as W  # noqa: E402
from openpyxl import Workbook  # noqa: E402
from openpyxl.drawing.image import Image as XLImage  # noqa: E402
from openpyxl.styles import Alignment  # noqa: E402
from openpyxl.utils import get_column_letter as CL  # noqa: E402
from excel_theme import write_audit_header, write_section_bar, write_table, plain_punct  # noqa: E402

TH, ACC, PCT = W.TH, W.ACC, "percent_acct"
GAAP = {"ifrs": "한국채택국제회계기준", "kgaap": "일반기업회계기준"}
TITLE = "실증적 분석절차"
PREC = ["예측가능성", "분리가능성", "자료 신뢰성", "절차 유형"]
LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)


def expected(t):
    """파이썬 기대치(기간별 · 항목별) - 조서 수식과 같은 계산."""
    d, p = t["driver"], t["param"]
    raw, fac = d["raw"]["rows"], (d.get("factor") or {}).get("value", 1)
    sh = d.get("share")
    div = (p.get("divisor") or {}).get("value", 1)
    out = []
    for i in range(len(t["periods"])):
        row = []
        for j, m in enumerate(d["map"]):
            q = raw[i][m["raw"]] * fac
            if "share" in m:
                b = sh["rows"][i]
                q *= b[m["share"]] / sum(b)
            row.append(q * p["rows"][i][j] / div)
        out.append(row)
    return out


def build(cfg, out):
    wb = Workbook()
    wb.remove(wb.active)
    wp_no = cfg.get("wp_no", "3300A")
    meta = lambda no, name=TITLE: {"회사명": cfg["company"], "조서번호": "%s-%s" % (wp_no, no) if no else wp_no,
                                   "결산일": cfg["year_end"], "조서명": name, "작성자": cfg.get("author", "")}
    s0 = wb.create_sheet(wp_no)
    pm = cfg["pm"]
    summ_rows = []
    sheets = []
    for n, t in enumerate(cfg["targets"], 1):
        ref = t.get("ref", "Ref.%d" % n)
        sm = wb.create_sheet(("%s %s" % (ref, t["account"]))[:31])
        sa = wb.create_sheet(("%s-A 드라이버" % ref)[:31])
        sb = wb.create_sheet(("%s-B 파라미터" % ref)[:31])
        sheets.append((ref, t, sm, sa, sb))

    # ── 요약(3300A) - 머리 · 공통 정보 ──
    hdr = write_audit_header(s0, meta(None), sheet_title=TITLE, theme=TH, tab="output")
    r = write_section_bar(s0, hdr, "1. 기본 정보", theme=TH, last_col=9)
    basic = [["회계기준", GAAP.get(cfg.get("gaap", "ifrs"))], ["수행중요성", pm["amount"]], ["출처", pm.get("source", "2700 중요성")]]
    write_table(s0, ["항목", "내용", "", ""], [x + [None, None] for x in basic], theme=TH, header_row=r, number_cols={},
                input_cells=["C%d:E%d" % (r + 2, r + 2)])
    W.merge_text_col(s0, r, r + 1, r + len(basic), 3, 5)
    s0["C%d" % (r + 2)].number_format = "#,###,##0;[Red](#,###,##0);-"
    s0["C%d" % (r + 2)].alignment = Alignment(horizontal="right", vertical="center")
    W.nm(wb, "수행중요성", wp_no, "$C$%d" % (r + 2))
    r = r + len(basic) + 2
    summ_r = write_section_bar(s0, r, "2. 실증적 분석절차 요약", theme=TH, last_col=9)

    res = {"targets": []}
    for ref, t, sm, sa, sb in sheets:
        P, I = t["periods"], [m["item"] for m in t["driver"]["map"]]
        nP, nI = len(P), len(I)
        d, p = t["driver"], t["param"]
        A, Bn, Mn = sa.title, sb.title, sm.title
        q = lambda s: "'%s'" % s

        # ── A 드라이버 ──
        h = write_audit_header(sa, meta("%s-A" % ref.split(".")[-1], "%s 드라이버" % t["account"]), theme=TH, tab="pbc")
        r = write_section_bar(sa, h, "1. 원자료", theme=TH, last_col=max(4, 2 + nI))
        rc = d["raw"]["columns"]
        rows = [[P[i]] + list(d["raw"]["rows"][i]) for i in range(nP)]
        f0 = r + 1
        rows.append(["합계"] + ["=SUM(%s%d:%s%d)" % (CL(3 + j), f0, CL(3 + j), f0 + nP - 1) for j in range(len(rc))])
        write_table(sa, ["기간"] + rc, rows, theme=TH, header_row=r, number_cols={j: ACC for j in range(1, len(rc) + 1)},
                    total_rows=[len(rows)], input_cells=["C%d:%s%d" % (f0, CL(2 + len(rc)), f0 + nP - 1)])
        raw_r = f0
        r = f0 + nP + 1
        W.put(sa, r, 2, "(Source : %s)" % d["raw"].get("source", ""))
        r += 2
        fac_ref = None
        if d.get("factor"):
            r = write_section_bar(sa, r, "2. 환산계수", theme=TH, last_col=max(4, 2 + nI))
            fx = d["factor"]
            la = max(5, 2 + nI)
            write_table(sa, ["파라미터", "값", "비고"] + [""] * (la - 4), [[fx["label"], fx["value"], fx.get("note")] + [None] * (la - 4)],
                        theme=TH, header_row=r, number_cols={1: "#,##0.####"}, input_cells=["C%d" % (r + 1)])
            W.merge_text_col(sa, r, r + 1, r + 1, 4, la)
            fac_ref = "%s!$C$%d" % (q(A), r + 1)
            r += 3
        sh_r = None
        if d.get("share"):
            sh = d["share"]
            r = write_section_bar(sa, r, "3. 안분 비중", theme=TH, last_col=max(4, 2 + nI))
            if sh.get("note"):
                W.put(sa, r - 1, 2, sh["note"])
                r += 1
            sc = sh["columns"]
            f0 = r + 1
            rows = []
            for i in range(nP):
                rr = f0 + i
                rows.append([P[i]] + list(sh["rows"][i]) +
                            ["=%s%d/SUM($C%d:$%s%d)" % (CL(3 + k), rr, rr, CL(2 + len(sc)), rr) for k in range(len(sc))])
            tot = f0 + nP
            rows.append(["합계"] + ["=SUM(%s%d:%s%d)" % (CL(3 + j), f0, CL(3 + j), tot - 1) for j in range(len(sc))] +
                        ["=%s%d/SUM($C%d:$%s%d)" % (CL(3 + k), tot, tot, CL(2 + len(sc)), tot) for k in range(len(sc))])
            write_table(sa, ["기간"] + sc + ["%s 비중" % c.split("(")[0] for c in sc], rows, theme=TH, header_row=r,
                        number_cols={**{j: ACC for j in range(1, len(sc) + 1)}, **{j: "0.00%" for j in range(len(sc) + 1, 2 * len(sc) + 1)}},
                        total_rows=[len(rows)], input_cells=["C%d:%s%d" % (f0, CL(2 + len(sc)), f0 + nP - 1)])
            sh_r = (f0, len(sc))
            r = tot + 1
            W.put(sa, r, 2, "(Source : %s)" % sh.get("source", ""))
            r += 2
        r = write_section_bar(sa, r, "4. %s(%s)" % (d["name"], d.get("unit", "")) if d.get("unit") else "4. %s" % d["name"], theme=TH,
                              last_col=max(4, 2 + nI))
        f0 = r + 1
        rows = []
        for i in range(nP):
            rr = f0 + i
            cells = []
            for m in d["map"]:
                e = "%s%d" % (CL(3 + m["raw"]), raw_r + i)
                if fac_ref:
                    e += "*" + fac_ref
                if "share" in m:
                    e += "*%s%d" % (CL(3 + sh_r[1] + m["share"]), sh_r[0] + i)
                cells.append("=" + e)
            rows.append([P[i]] + cells + ["=SUM(C%d:%s%d)" % (rr, CL(2 + nI), rr)])
        rows.append(["합계"] + ["=SUM(%s%d:%s%d)" % (CL(3 + j), f0, CL(3 + j), f0 + nP - 1) for j in range(nI + 1)])
        write_table(sa, ["기간"] + I + ["합계"], rows, theme=TH, header_row=r, number_cols={j: ACC for j in range(1, nI + 2)},
                    total_rows=[len(rows)])
        drv_r = f0
        sa.column_dimensions["B"].width = 22
        for j in range(2 * max(nI, len(rc)) + 2):
            sa.column_dimensions[CL(3 + j)].width = 16

        # ── B 파라미터 ──
        h = write_audit_header(sb, meta("%s-B" % ref.split(".")[-1], "%s 파라미터" % t["account"]), theme=TH, tab="pbc")
        r = write_section_bar(sb, h, "1. 외부 자료", theme=TH, last_col=max(4, 2 + nI))
        W.put(sb, r - 1, 2, "%s(%s)" % (p["name"], p.get("unit", "")))
        r += 1
        f0 = r + 1
        write_table(sb, ["기간"] + I, [[P[i]] + list(p["rows"][i]) for i in range(nP)], theme=TH, header_row=r,
                    number_cols={j: "#,##0.00" for j in range(1, nI + 1)}, input_cells=["C%d:%s%d" % (f0, CL(2 + nI), f0 + nP - 1)])
        par_r = f0
        r = f0 + nP
        for line in p.get("source_lines", ["(Source : %s)" % p.get("source", "")]):
            W.put(sb, r, 2, line)
            r += 1
        r += 1
        div_ref = None
        if p.get("divisor"):
            dv = p["divisor"]
            r = write_section_bar(sb, r, "2. 적용 단가", theme=TH, last_col=max(4, 2 + nI))
            lb = max(5, 2 + nI)
            write_table(sb, ["파라미터", "값", "비고"] + [""] * (lb - 4), [[dv["label"], dv["value"], dv.get("note")] + [None] * (lb - 4)],
                        theme=TH, header_row=r, number_cols={1: "0.00"}, input_cells=["C%d" % (r + 1)])
            W.merge_text_col(sb, r, r + 1, r + 1, 4, lb)
            div_ref = "%s!$C$%d" % (q(Bn), r + 1)
            r += 3
            f0 = r + 1
            write_table(sb, ["기간"] + I, [[P[i]] + ["=%s%d/%s" % (CL(3 + j), par_r + i, div_ref) for j in range(nI)] for i in range(nP)],
                        theme=TH, header_row=r, number_cols={j: "#,##0.00" for j in range(1, nI + 1)})
            par_r = f0
            r = f0 + nP + 1
        for cap in t.get("captures", []):
            if os.path.exists(cap["path"]):
                W.put(sb, r, 2, cap["label"], bold=True)
                img = XLImage(cap["path"])
                img.width, img.height = img.width * 0.6, img.height * 0.6
                sb.add_image(img, "B%d" % (r + 1))
                r += int(img.height / 20) + 3
        sb.column_dimensions["B"].width = 22
        for j in range(nI + 2):
            sb.column_dimensions[CL(3 + j)].width = 16

        # ── Ref 본 시트(9단계) ──
        h = write_audit_header(sm, meta(ref.split(".")[-1], "%s %s" % (TITLE, t["account"])), theme=TH, tab="output")
        last = 8
        r = write_section_bar(sm, h, "1. 대상·주장·범위", theme=TH, last_col=last)
        tb = t.get("tb_lines", [])
        rows = [[x[0], x[1], x[2] if len(x) > 2 else None] for x in tb]
        rows.append([t["account"] + "(대상)", None, t.get("target_note")])
        write_table(sm, ["계정", "금액", "비고"] + [""] * (last - 4), [x + [None] * (last - 4) for x in rows], theme=TH, header_row=r,
                    number_cols={1: ACC}, input_cells=["C%d:C%d" % (r + 1, r + len(tb))] if tb else [])
        W.merge_text_col(sm, r, r + 1, r + len(rows), 4, last)
        tgt_amt_cell = "C%d" % (r + len(rows))
        r = r + len(rows) + 1
        W.put(sm, r, 2, "경영진주장: %s" % t.get("assertion", ""))
        W.put(sm, r + 1, 2, t.get("scope", ""))
        r += 3
        r = write_section_bar(sm, r, "2. 정밀도 요인", theme=TH, last_col=last)
        pr = t.get("precision", {})
        write_table(sm, ["요인", "판단", "", "", "", ""], [[k, pr.get(k)] + [None] * 4 for k in PREC], theme=TH, header_row=r,
                    input_cells=["C%d:C%d" % (r + 1, r + len(PREC))])
        W.merge_text_col(sm, r, r + 1, r + len(PREC), 3, last)
        r = r + len(PREC) + 2
        r = write_section_bar(sm, r, "3. 기대치 모델", theme=TH, last_col=last)
        W.put(sm, r - 1, 2, t.get("model_text", ""))
        r += 1
        for line in t.get("key_factors", []):
            W.put(sm, r - 1, 2, line)
            r += 1
        r += 1
        r = write_section_bar(sm, r, "4. 자료 신뢰성", theme=TH, last_col=last)
        rel = t.get("reliability", [])
        write_table(sm, ["데이터", "출처", "", "절차", "", "결과", ""],
                    [[x[0], x[1], None, x[2], None, x[3] if len(x) > 3 else None, None] for x in rel],
                    theme=TH, header_row=r, input_cells=["G%d:H%d" % (r + 1, r + len(rel))] if rel else [])
        W.merge_text_col(sm, r, r + 1, r + len(rel), 3, 4)
        W.merge_text_col(sm, r, r + 1, r + len(rel), 5, 6)
        W.merge_text_col(sm, r, r + 1, r + len(rel), 7, 8)
        for i in range(len(rel)):
            sm["G%d" % (r + 1 + i)].alignment = LEFT
        r = r + len(rel) + 2
        r = write_section_bar(sm, r, "5. 수용가능한 차이", theme=TH, last_col=last)
        th_ = t.get("threshold", {})
        pad = [None] * (last - 4)
        write_table(sm, ["항목", "금액", "근거"] + [""] * (last - 4),
                    [["수행중요성(PM)", "=수행중요성", pm.get("source", "2700 중요성")] + pad,
                     ["수용가능한 차이 비율", th_.get("ratio", 1.0), th_.get("basis", "PM x 100%")] + pad,
                     ["수용가능한 차이(threshold)", "=C%d*C%d" % (r + 1, r + 2), th_.get("note", "PM x 비율")] + pad],
                    theme=TH, header_row=r, number_cols={1: ACC}, input_cells=["C%d" % (r + 2)])
        W.merge_text_col(sm, r, r + 1, r + 3, 4, last)
        sm["C%d" % (r + 2)].number_format = "0%"
        thr = "C%d" % (r + 3)
        r += 5
        if th_.get("text"):
            W.put(sm, r - 1, 2, th_["text"])
            r += 1
        r = write_section_bar(sm, r, "6. 기대치와 실제 비교", theme=TH, last_col=last)
        W.put(sm, r - 1, 2, "(1) 기대치 = %s × %s" % (d["name"], p["name"]))
        r += 1
        f0 = r + 1
        rows = []
        for i in range(nP):
            rr = f0 + i
            cells = ["=%s!%s%d*%s!%s%d" % (q(A), CL(3 + j), drv_r + i, q(Bn), CL(3 + j), par_r + i) for j in range(nI)]
            rows.append([P[i]] + cells + ["=SUM(C%d:%s%d)" % (rr, CL(2 + nI), rr)])
        rows.append(["합계"] + ["=SUM(%s%d:%s%d)" % (CL(3 + j), f0, CL(3 + j), f0 + nP - 1) for j in range(nI + 1)])
        write_table(sm, ["기간"] + I + ["합계"], rows, theme=TH, header_row=r, number_cols={j: ACC for j in range(1, nI + 2)},
                    total_rows=[len(rows)])
        exp_r, exp_col = f0, CL(3 + nI)
        r = f0 + nP + 2
        W.put(sm, r - 1, 2, "(2) 기대치(A)와 %s(B)" % t["actual"].get("label", "회사 기록 금액"))
        r += 1
        f0 = r + 1
        rows = []
        for i in range(nP):
            rr = f0 + i
            rows.append([P[i], "=%s%d" % (exp_col, exp_r + i), t["actual"]["rows"][i], "=D%d-C%d" % (rr, rr), '=IFERROR(E%d/D%d,"")' % (rr, rr)])
        tot = f0 + nP
        rows.append(["합계", "=SUM(C%d:C%d)" % (f0, tot - 1), "=SUM(D%d:D%d)" % (f0, tot - 1), "=D%d-C%d" % (tot, tot),
                     '=IFERROR(E%d/D%d,"")' % (tot, tot)])
        write_table(sm, ["기간", "기대치(A)", "실제(B)", "차이(B-A)", "차이율((B-A)/B)"], rows, theme=TH, header_row=r,
                    number_cols={1: ACC, 2: ACC, 3: ACC, 4: PCT}, total_rows=[len(rows)], input_cells=["D%d:D%d" % (f0, tot - 1)])
        sm[tgt_amt_cell] = "=D%d" % tot
        sm[tgt_amt_cell].number_format = "#,###,##0;[Red](#,###,##0);-"
        r = tot + 1
        W.put(sm, r, 2, "(Source : %s)" % t["actual"].get("source", ""))
        r += 2
        write_table(sm, ["구분", "결과"], [["수용가능한 차이", "=%s" % thr], ["총차이(절대값)", "=ABS(E%d)" % tot],
                                         ["판정", '=IF(ABS(E%d)<=%s,"수용가능한 차이 이내","수용가능한 차이 초과 - 차이 조사")' % (tot, thr)]],
                    theme=TH, header_row=r, number_cols={1: ACC}, secondary=True)
        sm["C%d" % (r + 3)].alignment = Alignment(horizontal="left", vertical="center")
        judge_cell, diff_cell, exp_cell = "C%d" % (r + 3), "E%d" % tot, "C%d" % tot
        r += 5
        for no, key, title in ((7, "investigation", "7. 차이 조사"), (8, "revision", "8. 수정 후 재비교"), (9, "conclusion", "9. 결론")):
            r = write_section_bar(sm, r, title, theme=TH, last_col=last)
            write_table(sm, [title[3:]] + [""] * (last - 2), [[t.get(key)] + [None] * (last - 2)], theme=TH, header_row=r,
                        input_cells=["B%d" % (r + 1)])
            W.merge_text_col(sm, r, r + 1, r + 1, 2, last)
            r += 3
        sm.column_dimensions["B"].width = 30
        sm.column_dimensions["C"].width = 24
        for j in range(nI + 2):
            sm.column_dimensions[CL(3 + j)].width = max(18, sm.column_dimensions[CL(3 + j)].width or 0)
        for c in "EFGH":
            sm.column_dimensions[c].width = max(18, sm.column_dimensions[c].width or 0)

        summ_rows.append((t, ref, Mn, exp_cell, thr, diff_cell, judge_cell))
        E = expected(t)
        res["targets"].append({"ref": ref, "account": t["account"], "기대치": sum(map(sum, E)), "실제": sum(t["actual"]["rows"]),
                               "차이": sum(t["actual"]["rows"]) - sum(map(sum, E)),
                               "threshold": pm["amount"] * th_.get("ratio", 1.0), "sheet": Mn, "exp_cell": exp_cell,
                               "diff_cell": diff_cell})

    # ── 요약 표 ──
    r = summ_r
    rows = []
    for t, ref, Mn, ec, thr, dc, jc in summ_rows:
        q = "'%s'" % Mn
        rows.append([t["account"], ref, t.get("model_text"), "=%s!%s" % (q, ec), "=%s!%s" % (q, thr), "=%s!%s" % (q, dc), "=%s!%s" % (q, jc),
                     t.get("investigation")])
    write_table(s0, ["관련 계정", "Ref.", "1. 기대치의 개발", "기대치", "2. 수용가능한 차이", "3. 차이 계산", "판정", "4. 유의적 차이의 조사"],
                rows, theme=TH, header_row=r, number_cols={3: ACC, 4: ACC, 5: ACC})
    for i, (t, ref, Mn, *_rest) in enumerate(summ_rows):
        rr = r + 1 + i
        W.link(s0["C%d" % rr], Mn)
        for c in "DHI":
            s0["%s%d" % (c, rr)].alignment = LEFT
    r = r + len(rows) + 2
    r = write_section_bar(s0, r, "3. 결론", theme=TH, last_col=9)
    write_table(s0, ["결론"] + [""] * 7, [[cfg.get("conclusion")] + [None] * 7], theme=TH, header_row=r, input_cells=["B%d" % (r + 1)])
    W.merge_text_col(s0, r, r + 1, r + 1, 2, 9)
    for c, w in (("B", 18), ("C", 10), ("D", 44), ("E", 18), ("F", 22), ("G", 18), ("H", 24), ("I", 40)):
        s0.column_dimensions[c].width = w
    # 상세 시트 → 요약으로 돌아가는 링크(머리 오른쪽)
    for ref, t, sm, sa, sb in sheets:
        for s in (sm, sa, sb):
            W.link(s.cell(2, 9), wp_no, text="%s 요약으로" % wp_no)
    W.fit_headers(wb)
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("input")
    ap.add_argument("--out", default="")
    ap.add_argument("--overwrite", action="store_true")
    ap.add_argument("--no-fit", action="store_true")
    a = ap.parse_args()
    cfg = json.load(open(a.input, encoding="utf-8"))
    out = a.out or cfg["output"]
    if os.path.exists(out) and not a.overwrite:
        raise KicpaError("이미 있습니다(사용자가 고쳤을 수 있음 - 확인 후 --overwrite 또는 다른 --out): %s" % out)
    for t in cfg["targets"]:
        nP = len(t["periods"])
        for k, rows in (("driver.raw", t["driver"]["raw"]["rows"]), ("param", t["param"]["rows"]), ("actual", t["actual"]["rows"])):
            if len(rows) != nP:
                raise KicpaError("%s 의 %s 줄 수(%d)가 기간 수(%d)와 다릅니다" % (t["account"], k, len(rows), nP))
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    res = build(cfg, out)
    json.dump(res, open(os.path.splitext(out)[0] + "_summary.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    if not a.no_fit:
        print(W.fit_rows(out))
    print("저장: %s" % out)
    for x in res["targets"]:
        print("  %s %s 기대치 %s · 실제 %s · 차이 %s · 수용가능한 차이 %s → %s" % (
            x["ref"], x["account"], "{:,.0f}".format(x["기대치"]), "{:,.0f}".format(x["실제"]), "{:,.0f}".format(x["차이"]),
            "{:,.0f}".format(x["threshold"]), "이내" if abs(x["차이"]) <= x["threshold"] else "초과 - 차이 조사"))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        main()
    except KicpaError as ex:
        print("ERROR\t%s" % ex)
        sys.exit(2)
