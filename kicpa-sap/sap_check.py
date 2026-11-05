# -*- coding: utf-8 -*-
"""kicpa-sap 검증 - 3300A 조서를 LibreOffice 로 재계산해 대상별 기대치 · 차이가 파이썬 계산(<조서>_summary.json)과 같은지 본다.

  python -X utf8 sap_check.py <3300A 조서.xlsx> [--expect <기대치> <차이>]   (--expect: 기존 조서 값과도 대조)
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
    exp = [float(x) for x in sys.argv[sys.argv.index("--expect") + 1:][:2]] if "--expect" in sys.argv else None
    S = json.load(open(os.path.splitext(xl)[0] + "_summary.json", encoding="utf-8"))
    path, tmp = recalc(xl)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        bad = []
        for t in S["targets"]:
            ws = wb[t["sheet"]]
            e, d = ws[t["exp_cell"]].value, ws[t["diff_cell"]].value
            ok = isinstance(e, (int, float)) and abs(e - t["기대치"]) < 1 and abs(d - t["차이"]) < 1
            print("  %s %s 엑셀 기대치 %s · 차이 %s / 파이썬 %s · %s %s" % (
                t["ref"], t["account"], "{:,.0f}".format(e or 0), "{:,.0f}".format(d or 0), "{:,.0f}".format(t["기대치"]),
                "{:,.0f}".format(t["차이"]), "OK" if ok else "<< 다름"))
            if not ok:
                bad.append(t["ref"])
            if exp and len(S["targets"]) == 1:
                ok2 = abs(e - exp[0]) < 1 and abs(d - exp[1]) < 1
                print("  기존 조서 값 기대치 %s · 차이 %s %s" % ("{:,.0f}".format(exp[0]), "{:,.0f}".format(exp[1]), "OK" if ok2 else "<< 다름"))
                if not ok2:
                    bad.append("기존 조서")
        errs = error_cells(wb)
        print("  오류 값 칸 %d%s" % (len(errs), (" 예: " + ", ".join(errs[:5])) if errs else ""))
        if errs:
            bad.append("오류 값")
        print("결과: " + ("모두 일치" if not bad else "다름 - %s" % bad))
        sys.exit(1 if bad else 0)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
