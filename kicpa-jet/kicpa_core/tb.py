# -*- coding: utf-8 -*-
"""TB 읽기 - 금액은 모두 차변 +(대변 잔액은 음수).

세 형식을 받는다.
  표(xlsx · csv)    머리글: 계정코드 · 계정과목 · [구분] · [표시과목] · 기초(전기말) · 기말(당기)
                    또는 차변잔액 · 대변잔액(기말만) - 전기말 표를 따로 주면 --tb-prior 로
  xlaudit tb 폴더   coa.csv(계층) + tb.csv(기준일 · 코드 · 금액) - 기준일 두 개를 고른다
  douzone_tb json   {"rows": [[코드, 계정과목, 당기, 전기말]], "unap": [당기, 전기]} (xlaudit douzone_tb.py 산출)
반환 행: {코드, 계정과목, 구분, 표시과목, 기초, 기말}
"""
import csv
import json
import os
import re

from . import KicpaError
from .journal import code_of, load_rows, norm, to_num

TBHEAD = {"코드": ("계정코드", "코드", "계정과목코드"), "계정과목": ("계정과목", "계정명", "계정과목명", "이름"),
          "구분": ("구분", "대분류", "분류"), "표시과목": ("표시과목", "공시과목", "공시", "재무제표과목"),
          "기초": ("기초", "전기말", "전기", "기초잔액", "전기말잔액"), "기말": ("기말", "당기", "당기말", "기말잔액", "잔액", "금액"),
          "차변잔액": ("차변잔액",), "대변잔액": ("대변잔액",)}


def _table(path):
    rows = load_rows(path)
    for i, r in enumerate(rows[:40]):
        cells = [norm(x) for x in r]
        if any(c in TBHEAD["코드"] for c in cells) and (any(c in TBHEAD["기말"] for c in cells) or "차변잔액" in cells):
            col = {}
            for j, c in enumerate(cells):
                for f, names in TBHEAD.items():
                    if f not in col and c in names:
                        col[f] = j
            out = []
            for r2 in rows[i + 1:]:
                g = lambda f: r2[col[f]] if f in col and col[f] < len(r2) else None
                code = code_of(g("코드"))
                if code in (None, ""):
                    continue
                if "기말" in col:
                    cur = to_num(g("기말"))
                else:
                    cur = to_num(g("차변잔액")) - to_num(g("대변잔액"))
                out.append({"코드": code, "계정과목": str(g("계정과목") or "").strip(), "구분": str(g("구분") or "").strip() or None,
                            "표시과목": str(g("표시과목") or "").strip() or None,
                            "기초": to_num(g("기초")) if "기초" in col else None, "기말": cur})
            return out
    raise KicpaError("TB 머리글(계정코드 · 기말)을 찾지 못했습니다: %s" % os.path.basename(path))


def from_table(path, prior=None):
    rows = _table(path)
    if prior:
        p = {r["코드"]: r["기말"] for r in _table(prior)}
        for r in rows:
            r["기초"] = p.pop(r["코드"], 0.0)
        for c, v in p.items():
            rows.append({"코드": c, "계정과목": "", "구분": None, "표시과목": None, "기초": v, "기말": 0.0})
    for r in rows:
        if r["기초"] is None:
            r["기초"] = 0.0
    return rows, {"형식": "TB 표", "파일": os.path.basename(path)}


def from_xlaudit(d, cur, prior):
    """xlaudit tb 폴더(coa.csv · tb.csv) → 기준일 cur · prior 두 열. 구분 = 대분류(lv1), 표시과목 = 공시(lv3)."""
    coa = {r["코드"]: r for r in csv.DictReader(open(os.path.join(d, "coa.csv"), encoding="utf-8-sig"))}
    dates = set()
    amt = {}
    for r in csv.DictReader(open(os.path.join(d, "tb.csv"), encoding="utf-8-sig")):
        dates.add(r["기준일"])
        amt[(r["기준일"], r["코드"])] = to_num(r["금액"])
    for x in (cur, prior):
        if x not in dates:
            raise KicpaError("TB 기준일 %s 이 없습니다(있는 기준일: %s)" % (x, ", ".join(sorted(dates))))

    def up(c, lv):
        while c in coa and coa[c]["Level"] != lv:
            c = coa[c]["상위코드"]
        return coa[c]["이름"] if c in coa else None

    rows = []
    for c, r in coa.items():
        if r["Level"] != "lv5":
            continue
        a, b = amt.get((cur, c), 0.0), amt.get((prior, c), 0.0)
        if not a and not b:
            continue
        rows.append({"코드": code_of(c), "계정과목": r["이름"], "구분": up(c, "lv1"), "표시과목": up(c, "lv3"),
                     "기초": b, "기말": a})
    return rows, {"형식": "xlaudit tb 폴더", "파일": os.path.basename(os.path.normpath(d)), "당기일": cur, "전기일": prior}


def from_douzone_json(path):
    d = json.load(open(path, encoding="utf-8"))
    rows = [{"코드": code_of(c), "계정과목": n, "구분": None, "표시과목": None, "기초": p, "기말": a}
            for c, n, a, p in d["rows"] if code_of(c) not in (None, "")]
    return rows, {"형식": "douzone_tb json", "파일": os.path.basename(path), "unap": d.get("unap")}


def load(tb=None, tb_dir=None, tb_json=None, cur=None, prior=None, tb_prior=None):
    if tb_dir:
        if not (cur and prior):
            raise KicpaError("xlaudit tb 폴더는 --tb-date(당기) · --tb-prior-date(전기말) 를 함께 줍니다")
        return from_xlaudit(tb_dir, cur, prior)
    if tb_json:
        return from_douzone_json(tb_json)
    if tb:
        return from_table(tb, tb_prior)
    raise KicpaError("TB 를 주지 않았습니다(--tb · --tb-dir · --tb-json 중 하나)")


def date_key(s):
    """'2026.06.30' · '2026-06-30' · '20260630' → date."""
    import datetime as dt
    m = re.match(r"^(\d{4})[-./]?(\d{1,2})[-./]?(\d{1,2})$", str(s).strip())
    if not m:
        raise KicpaError("날짜 형식이 아닙니다: %s" % s)
    return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
