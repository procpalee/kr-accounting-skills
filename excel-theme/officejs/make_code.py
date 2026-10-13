# -*- coding: utf-8 -*-
"""
make_code.py - excel-theme 서식을 "그대로 실행할 Office.js 코드"로 출력한다 (표준 라이브러리만 사용).

사용법:
  python make_code.py themes
  python make_code.py table '<JSON>'      # 표 서식
  python make_code.py parts '<JSON>'      # 제목 밴드 · 감사조서 헤더 · 섹션 바 · 캡션 · 탭 색
  (JSON 대신 @파일경로 또는 표준입력도 가능)

공통 JSON 키:
  theme       : default | audit | procpa | dcf-valuation   (생략 시 default)
  sheet_name  : 대상 시트명 (생략 시 활성 시트)
  body_only   : true 면 Excel.run 래퍼 없이 본문만 출력 (실행 도구가 context 를 이미 줄 때)
  json        : true 면 {code, ops, next_header_row, warnings} JSON 출력

table 키:
  n_cols, n_rows (필수) · header_row(4) · start_col(2=B) · title_cell("auto"=B2, audit 은 없음 · ""=없음)
  currency_cols / percent_cols : 0-based 데이터열 번호 목록
  number_format_cols : {"C":"accounting", "1":"multiple"}  열문자 또는 0-based 번호 → 서식키
  section_rows / subtotal_rows / total_rows : 데이터 1-based 행 번호 목록
  input_cells / linked_cells / todo_cells   : ["C5:C6"] 형식 범위 목록
  secondary(false) : 서브표 색 · autofit(true) : 열너비 자동(상한 30)

parts 키:
  parts : [{"type":"title_band","last_col":7},
           {"type":"audit_header","meta":{"회사명":"","조서번호":"","결산일":"","조서명":"","작성자":""}},
           {"type":"section_bar","row":6,"text":"1. 총괄표","last_col":8},
           {"type":"caption","cell":"B12","text":"표1. 재질별 단가"},
           {"type":"tab","role":"calc"}]
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from officejs_plan import officejs_plan, officejs_sheet_parts  # noqa: E402

TABLE_KEYS = {"n_cols", "n_rows", "header_row", "start_col", "title_cell", "number_format_cols",
              "currency_cols", "percent_cols", "section_rows", "subtotal_rows", "total_rows",
              "input_cells", "linked_cells", "todo_cells", "secondary", "autofit"}


def _load():
    with open(os.path.join(HERE, "theme.json"), encoding="utf-8") as f:
        return json.load(f)


def _theme(data, name):
    name = (name or data["default"]).strip()
    name = data["aliases"].get(name, name)
    if name not in data["themes"]:
        raise SystemExit("알 수 없는 테마 '%s'. 사용 가능: %s" % (name, ", ".join(data["themes"])))
    return name


def _read_arg(argv):
    raw = argv[2] if len(argv) > 2 else sys.stdin.read()
    raw = raw.strip()
    if raw.startswith("@"):
        with open(raw[1:], encoding="utf-8") as f:
            raw = f.read()
    try:
        return json.loads(raw or "{}")
    except json.JSONDecodeError as e:
        raise SystemExit("JSON 형식 오류: %s" % e)


def main(argv):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    if len(argv) < 2 or argv[1] not in ("themes", "table", "parts"):
        print(__doc__)
        return 2
    data = _load()
    if argv[1] == "themes":
        print(json.dumps({"default": data["default"],
                          "themes": {k: v["label"] for k, v in data["themes"].items()},
                          "aliases": data["aliases"],
                          "tab_roles": ["guide", "output", "calc", "input", "pbc", "raw"]},
                         ensure_ascii=False, indent=1))
        return 0

    args = _read_arg(argv)
    theme = _theme(data, args.pop("theme", ""))
    spec, rules = data["themes"][theme], data.get("rules")
    sheet = args.pop("sheet_name", "")
    body_only = bool(args.pop("body_only", False))
    as_json = bool(args.pop("json", False))

    if argv[1] == "table":
        unknown = set(args) - TABLE_KEYS
        if unknown:
            raise SystemExit("table 에 없는 키: %s (사용 가능: %s)" % (sorted(unknown), sorted(TABLE_KEYS)))
        if "n_cols" not in args or "n_rows" not in args:
            raise SystemExit("table 에는 n_cols 와 n_rows 가 필요하다.")
        nfc = args.get("number_format_cols") or {}
        args["number_format_cols"] = {(int(k) if str(k).isdigit() else k): v for k, v in nfc.items()}
        res = officejs_plan(spec, sheet_name=sheet, rules=rules, body_only=body_only, **args)
        res.setdefault("next_header_row", None)
        res.setdefault("warnings", [])
    else:
        parts = args.pop("parts", None)
        if args:
            raise SystemExit("parts 에 없는 키: %s" % sorted(args))
        if not parts:
            raise SystemExit('parts 목록이 필요하다. 예: {"parts":[{"type":"title_band","last_col":7}]}')
        res = officejs_sheet_parts(spec, parts, sheet_name=sheet, rules=rules, body_only=body_only)

    if as_json:
        print(json.dumps(res, ensure_ascii=False, indent=1))
        return 0
    head = ["// excel-theme %s (%s) - 아래 코드를 수정 없이 그대로 실행한다." % (argv[1], theme)]
    if res.get("next_header_row"):
        head.append("// next_header_row: %d  (다음 표의 header_row)" % res["next_header_row"])
    for w in res.get("warnings") or []:
        head.append("// WARNING: %s" % w)
    print("\n".join(head + [res["code"]]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
