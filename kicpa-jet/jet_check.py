# -*- coding: utf-8 -*-
"""kicpa-jet 검증 - 조서를 LibreOffice(헤드리스)로 재계산해 수식 결과가 파이썬 판정(<조서>_summary.json)과 같은지 본다.

  python -X utf8 jet_check.py <3600D 조서.xlsx> [<summary.json>]

대조: 시나리오별 추출 전표 수 · 줄 수(30 선정기준) · 분개장 차변 합계 · 완전성 차이 계정(20 모집단 완전성) · 오류 값 칸.
원본 조서는 읽기만 하고, 재계산 사본은 임시 폴더에 만든 뒤 지운다.
"""
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from kicpa_core.recalc import recalc, error_cells  # noqa: E402

import openpyxl  # noqa: E402


def labels(ws, col=2):
    out = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, col).value
        if isinstance(v, str):
            out.setdefault(v, r)
    return out


def main():
    xl = sys.argv[1]
    sj = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(xl)[0] + "_summary.json"
    S = json.load(open(sj, encoding="utf-8"))
    path, tmp = recalc(xl)
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        bad = []
        s30 = wb["30 선정기준"]
        lab = labels(s30)
        hdr = lab["시나리오 ID"]
        rv, rl = lab["추출 전표 수"], lab["추출 줄 수"]
        for j in range(3, s30.max_column + 1):
            s = s30.cell(hdr, j).value
            if s not in S["시나리오"]:
                continue
            got = (s30.cell(rv, j).value, s30.cell(rl, j).value)
            exp = (S["시나리오"][s]["전표"], S["시나리오"][s]["줄"])
            ok = tuple(int(round(x or 0)) for x in got) == tuple(exp)
            print("  %s 엑셀 전표 %s · 줄 %s / 파이썬 %s · %s %s" % (s, got[0], got[1], exp[0], exp[1], "OK" if ok else "<< 다름"))
            if not ok:
                bad.append(s)
        s20 = wb["20 모집단 완전성"]
        lab = labels(s20)
        for k in ("분개장 차변 - 대변", "차대가 다른 전표 수", "원본 합계 줄과 차이", "차이가 있는 계정 수", "차이 사유가 빈 계정 수"):
            print("  20 %s = %s" % (k, s20.cell(lab[k], 3).value))
        tot = s20.cell(lab["합계"], 5).value
        if abs((tot or 0) - S["차변"]) > 0.5:
            bad.append("차변 합계")
        print("  20 차변 합계 엑셀 %s / 파이썬 %s" % ("{:,.0f}".format(tot or 0), "{:,.0f}".format(S["차변"])))
        h = next(r for r in range(1, s20.max_row + 1) if s20.cell(r, 11).value == "차이 사유")
        got = {}
        for r in range(h + 1, s20.max_row + 1):
            code = s20.cell(r, 2).value
            if code is None and s20.cell(r, 3).value is None:
                break
            d = s20.cell(r, 10).value
            if isinstance(d, (int, float)) and abs(d) > 0.5:
                got[str(code)] = d
        exp = S.get("완전성 차이", {})
        keys = sorted(set(got) | set(exp), key=str)
        print("  완전성 차이 계정 %d개(파이썬 %d개)" % (len(got), len(exp)))
        for k in keys:
            a, b = got.get(k), exp.get(k)
            same = a is not None and b is not None and abs(a - b) <= 1
            if not same:
                bad.append("완전성 " + k)
            print("    %s 엑셀 %s / 파이썬 %s %s" % (k, "{:,.0f}".format(a) if a is not None else "-",
                                               "{:,.0f}".format(b) if b is not None else "-", "OK" if same else "<< 다름"))
        errs = error_cells(wb)
        print("  오류 값(#REF! 등) 칸 수: %d%s" % (len(errs), (" 예: " + ", ".join(errs[:5])) if errs else ""))
        if errs:
            bad.append("오류 값")
        print("결과: " + ("모두 일치" if not bad else "다름 - %s" % bad))
        sys.exit(1 if bad else 0)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
