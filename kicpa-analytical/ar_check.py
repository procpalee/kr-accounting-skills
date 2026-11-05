# -*- coding: utf-8 -*-
"""kicpa-analytical 검증 - 조서를 LibreOffice 로 재계산해 과목 금액 · 유의적 차이(잠정) · 검증 줄이 파이썬 계산과 같은지 본다.

  python -X utf8 ar_check.py <2120A · 8110A 조서.xlsx>      (같은 이름 _work\\summary.json 을 읽음)
"""
import json
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "kicpa-jet"))
from kicpa_core.recalc import recalc, error_cells  # noqa: E402

import openpyxl  # noqa: E402


def main():
    xl = sys.argv[1]
    S = json.load(open(os.path.join(os.path.splitext(xl)[0] + "_work", "summary.json"), encoding="utf-8"))
    path, tmp = recalc(xl)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet = next(n for n in wb.sheetnames if n in ("2120A", "8110A") or n == S.get("wp"))
        ws = wb[sheet]
        bad, n = [], 0
        for ln in S["lines"]:
            key = "%s:%s" % (ln["구역"], ln["키"])
            if key not in S["rowof"]:
                continue                                # 제조원가 계정이 없어 조서에 없는 구역
            r = S["rowof"][key]
            for col, k in (("C", "당기"), ("E", "전기")):
                v = ws["%s%d" % (col, r)].value
                if not isinstance(v, (int, float)) or abs(v - ln[k]) > 1:
                    bad.append("%s %s %s 엑셀 %s / 파이썬 %s" % (ln["과목"], k, col, v, round(ln[k])))
            flag = ws["I%d" % r].value or ""
            if ln["잠정"] and flag != ln["잠정"]:
                bad.append("%s 잠정 엑셀 %s / 파이썬 %s" % (ln["과목"], flag, ln["잠정"]))
            n += 1
        r0 = S["검증 행"]
        for i in range(4):
            lab = ws["B%d" % (r0 + i)].value
            a, b = ws["C%d" % (r0 + i)].value, ws["D%d" % (r0 + i)].value
            ok = all(x in (None, "") or (isinstance(x, (int, float)) and abs(x) < 1) for x in (a, b))
            print("  검증 %s: 당기 %s · 전기 %s %s" % (lab, a, b, "OK" if ok else "<< 확인"))
            if not ok:
                bad.append("검증 " + str(lab))
        rs = wb[next(x for x in wb.sheetnames if x.endswith("재무비율"))]
        shown = [(rs.cell(r, 3).value, rs.cell(r, 5).value, rs.cell(r, 6).value) for r in range(1, rs.max_row + 1)
                 if isinstance(rs.cell(r, 3).value, str) and rs.cell(r, 2).value in ("안정성", "수익성", "활동성", "성장성")]
        print("  재무비율 %d개 - 예: %s" % (len(shown), ", ".join("%s %s" % (a, round(b, 4) if isinstance(b, float) else b)
                                                         for a, b, c in shown[:4])))
        errs = error_cells(wb)
        print("  과목 %d줄 대조 · 다른 것 %d · 오류 값 칸 %d%s" % (n, len(bad), len(errs), (" 예: " + ", ".join(errs[:5])) if errs else ""))
        for b in bad[:20]:
            print("    " + b)
        if errs:
            bad.append("오류 값")
        print("결과: " + ("모두 일치" if not bad else "다름"))
        sys.exit(1 if bad else 0)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
