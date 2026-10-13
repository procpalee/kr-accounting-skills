# -*- coding: utf-8 -*-
"""
check_officejs.py — 스킬 패키지 officejs_plan 산출 코드 검증 (Excel 불필요).

1) 전 테마 × 대표 레이아웃으로 code 생성 → 같은 입력 2회 결과가 동일한지(결정론)
2) Node 로 가짜 Excel(Office.js 목) 위에서 code 를 실제 실행 → 문법·호출 오류 없음 확인
3) 행높이 op 가 없는지(사용자 규칙: 행높이 지정 금지)

    python officejs/tests/check_officejs.py
"""
import json
import os
import subprocess
import sys
import tempfile

SKILL = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # excel-theme
ROOT = os.path.join(SKILL, "dist", "excel-theme-officejs", "scripts")   # 빌드된 스킬 패키지
sys.path.insert(0, ROOT)
from officejs_plan import officejs_plan, officejs_sheet_parts  # noqa: E402

MOCK = r"""
const log = [];
function node(path) {
  return new Proxy(function () {}, {
    get(_, k) {
      if (k === "then") return undefined;
      if (k === Symbol.toPrimitive) return () => 0;
      return node(path + "." + String(k));
    },
    set(_, k, v) { log.push(["set", path + "." + String(k), v]); return true; },
    apply(_, __, args) { log.push(["call", path, args]); return node(path + "()"); },
  });
}
globalThis.Excel = {
  ConditionalFormatType: { cellValue: "CellValue" },
  run: async (fn) => fn({ workbook: node("workbook"), sync: async () => log.push(["sync"]) }),
};
const code = require("fs").readFileSync(process.argv[2], "utf8");
const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
new AsyncFunction(code)().then(() => {
  console.log(JSON.stringify({ ok: true, events: log.length }));
}).catch((e) => { console.log(JSON.stringify({ ok: false, error: String(e) })); process.exit(1); });
"""

LAYOUTS = [
    dict(n_cols=4, n_rows=6, currency_cols=[1, 2], percent_cols=[3], section_rows=[1], subtotal_rows=[5],
         total_rows=[6], input_cells=["C5:C6"], linked_cells=["D5:D6"], todo_cells=["E7"]),
    dict(n_cols=1, n_rows=1, section_rows=[1], title_cell=""),        # 1×1 엣지케이스
    dict(header_row=6, start_col=3, n_cols=8, n_rows=30, number_format_cols={"D": "multiple"},
         section_rows=[1, 10], total_rows=[29], subtotal_rows=[30], secondary=True, autofit=False),
]
PARTS = {
    "audit": [{"type": "audit_header", "meta": {"회사명": "A사", "조서번호": "6000A-10"}},
              {"type": "section_bar", "row": 6, "text": "1. 총괄표"},
              {"type": "caption", "cell": "B8", "text": "표1. 잔액"}, {"type": "tab", "role": "output"}],
    "*": [{"type": "title_band", "last_col": 9}, {"type": "caption", "cell": "B4", "text": "표1. 잔액"},
          {"type": "tab", "role": "guide"}],
}


def run_node(mock, code):
    f = tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8")
    f.write(code)
    f.close()
    p = subprocess.run(["node", mock, f.name], capture_output=True, text=True, encoding="utf-8")
    os.unlink(f.name)
    return json.loads(p.stdout.strip() or '{"ok": false, "error": %s}' % json.dumps(p.stderr))


def main():
    with open(os.path.join(ROOT, "theme.json"), encoding="utf-8") as f:
        data = json.load(f)
    themes, rules = data["themes"], data.get("rules")
    mock = tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8")
    mock.write(MOCK)
    mock.close()
    fails = 0
    for name, spec in themes.items():
        jobs = [("layout%d" % i, lambda lay=lay, i=i: officejs_plan(
                    spec, sheet_name="표 1" if i == 2 else "", rules=rules, **lay))
                for i, lay in enumerate(LAYOUTS)]
        jobs.append(("parts", lambda: officejs_sheet_parts(spec, PARTS.get(name, PARTS["*"]), rules=rules)))
        for tag, make in jobs:
            a, b = make(), make()
            det = json.dumps(a, sort_keys=True) == json.dumps(b, sort_keys=True)
            no_height = not any(o["op"] in ("autofitRows", "rowHeight") for o in a["ops"])
            res = run_node(mock.name, a["code"])
            ok = det and no_height and res.get("ok")
            fails += 0 if ok else 1
            print("%-4s %-14s %-8s ops=%-3d lines=%-3d det=%s noRowHeight=%s run=%s %s" % (
                "OK" if ok else "FAIL", name, tag, len(a["ops"]), a["code"].count("\n") + 1,
                det, no_height, res.get("ok"), "" if ok else res.get("error", "")))
    os.unlink(mock.name)
    print("\n%s (%d fail)" % ("ALL PASS" if not fails else "FAILED", fails))
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
