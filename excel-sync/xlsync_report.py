# -*- coding: utf-8 -*-
"""xlsync_report - diff 결과에 '고칠 코드 위치'를 붙이고 사람이 읽는 Markdown 으로 만든다.

코드 위치는 두 가지로 찾는다(둘 다 결정론):
  set_at  : 생성 기록(prov)에서 그 칸·속성을 마지막으로 쓴 코드 줄(파일:줄 함수 <- 호출한 곳).
  literal : 코드 폴더에서 바뀌기 전 문구·수식 안 문자열·색 코드가 그대로 적힌 줄.
"""
import json
import os

from xlsync_refs import expand, literals, rc

KIND_KO = {"value": "값", "style": "서식", "width": "열너비", "height": "행높이", "merge": "병합", "dv": "데이터 유효성",
           "cf": "조건부 서식", "comment": "메모", "link": "링크", "checkbox": "체크박스", "drawing": "그림·차트", "binary": "이진 파트",
           "sheet": "시트 속성", "row_insert": "행 삽입", "row_delete": "행 삭제", "col_insert": "열 삽입", "col_delete": "열 삭제"}
SKIP_DIRS = {"__pycache__", "_backup", "_deprecated", "_xlsync", ".git", "node_modules", "dist"}
EXTS = (".py", ".ps1", ".psm1", ".bas", ".vba", ".js", ".ts", ".json", ".yaml", ".yml", ".toml", ".ini")   # 설명서·데이터 덤프(.md·.txt·.csv)는 제외


def code_index(roots, max_bytes=300_000):
    idx = []
    for root in roots or []:
        if os.path.isfile(root):
            files = [root]
        else:
            files = []
            for d, dirs, fs in os.walk(root):
                dirs[:] = sorted(x for x in dirs if x not in SKIP_DIRS and not x.startswith("."))
                files += [os.path.join(d, f) for f in sorted(fs) if f.lower().endswith(EXTS)]
        for p in files:
            try:
                if os.path.getsize(p) > max_bytes:
                    continue
                with open(p, encoding="utf-8-sig") as f:
                    idx.append((p, f.read().splitlines()))
            except (UnicodeDecodeError, OSError):
                continue
    return idx


def find(idx, text, base, limit=4):
    if not isinstance(text, str) or len(text.strip()) < 2:
        return []
    variants = {text, json.dumps(text, ensure_ascii=False)[1:-1]}
    if "\n" in text:
        variants |= {text.split("\n")[0]}
    hits = []
    for p, lines in idx:
        for i, ln in enumerate(lines, 1):
            if len(ln) <= 1000 and any(v and v in ln for v in variants):      # 한 줄짜리 JSON 덤프는 건너뜀
                hits.append("%s:%d" % (_rel(p, base), i))
                if len(hits) >= limit:
                    return hits
    return hits


def _rel(p, base):
    try:
        r = os.path.relpath(p, base)
        return (p if r.startswith("..") else r).replace("\\", "/")
    except ValueError:
        return p.replace("\\", "/")


_ATTR = {"font": "font", "fill": "fill", "border": "border", "align": "alignment", "prot": "protection", "nf": "number_format"}


def _sites(prov, change):
    if not prov:
        return []
    sh = prov["sheets"].get(change.get("base_sheet") or change["sheet"], {})
    k, got = change["kind"], []
    if k in ("value", "style", "comment", "link", "checkbox"):
        want = {"value": ["value"], "comment": ["comment"], "link": ["hyperlink"], "checkbox": ["value"]}.get(k)
        if k == "style":
            a = change.get("attr", "").split(".")[0]
            want = [_ATTR.get(a, a), "style"]
        for rng in change.get("base_range", []):
            for r, c in (expand(rng) or [])[:400]:
                from xlsync_refs import a1
                attrs = sh.get("cells", {}).get(a1(r, c), {})
                for w in want + (["value"] if k == "style" else []):      # 서식을 따로 쓴 적이 없으면 그 칸을 만든 줄
                    if w in attrs:
                        got.append(attrs[w])
                        break
    elif k in ("width", "height"):
        d = sh.get("cols" if k == "width" else "rows", {}).get(str(change.get("base_at")), {})
        a = change.get("attr")
        if a in d:
            got.append(d[a])
    elif k == "merge":
        for rng in change.get("range", []):
            if rng in sh.get("merges", {}):
                got.append(sh["merges"][rng])
    elif k == "dv":
        cells = {rc2 for rng in change.get("base_range", []) for rc2 in (expand(rng) or [])}
        for d in sh.get("dv", []):
            if cells & set(expand(d["sqref"]) or []):
                got += [s for _, s in d.get("cells", [])] + ([d["made"]] if d.get("made") is not None else [])
    seen, out = set(), []
    for s in got:
        if s not in seen:
            seen.add(s)
            out.append(prov["sites"][s])
    return out[:3]


def attach(result, prov=None, idx=None, base="."):
    for c in result["changes"]:
        code = {"set_at": _sites(prov, c)}
        texts = []
        b = c.get("before")
        if c["kind"] == "value":
            if isinstance(b, str):
                texts.append(b)
            elif isinstance(b, dict) and "f" in b:
                texts += [t for t in literals(b["f"]) if len(t) >= 3]
        elif c["kind"] == "style" and isinstance(b, str) and len(b) == 6:
            texts.append(b)
        elif c["kind"] in ("row_delete", "col_delete"):
            texts += [p.split("=", 1)[1] for p in c.get("preview", [])[:2] if "=" in p]
        hits = []
        for t in texts[:3]:
            hits += [h for h in find(idx or [], t, base) if h not in hits]
        code["literal"] = hits[:4]
        c["code"] = code
    for n in result.get("names", []):
        n["code"] = {"set_at": [prov["sites"][prov["names"][n["name"]]]] if prov and n["name"] in prov.get("names", {}) else []}
    return result


def _fmt(v, width=60):
    if isinstance(v, dict) and "f" in v:
        s = v["f"]
    elif isinstance(v, (dict, list)):
        s = json.dumps(v, ensure_ascii=False, sort_keys=True)
    elif v is None:
        s = "(없음)"
    else:
        s = str(v)
    s = s.replace("|", "¦").replace("\n", "⏎")
    return s if len(s) <= width else s[:width - 1] + "…"


def _raw(v):
    if isinstance(v, dict) and "f" in v:
        return v["f"]
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False, sort_keys=True)
    return "(없음)" if v is None else str(v)


def _pair(b, a, width=70):
    """긴 문구·수식은 달라진 부분 앞뒤만 보여 준다."""
    sb, sa = _raw(b), _raw(a)
    if len(sb) <= width and len(sa) <= width:
        return _fmt(b, width), _fmt(a, width)
    p = 0
    while p < min(len(sb), len(sa)) and sb[p] == sa[p]:
        p += 1
    s = 0
    while s < min(len(sb), len(sa)) - p and sb[-1 - s] == sa[-1 - s]:
        s += 1

    def cut(x):
        lo, hi = max(0, p - 20), len(x) - s + 15
        return _fmt(("…" if lo > 0 else "") + x[lo:hi] + ("…" if hi < len(x) else ""), 110)
    return cut(sb), cut(sa)


def markdown(result, title="엑셀 변경 사항"):
    s = result["summary"]
    lines = ["# %s - %s" % (title, os.path.basename(result["cur"].get("file", ""))), "",
             "- 기준: %s %s · 현재: %s" % (result["base"].get("label") or "", result["base"].get("mtime", ""), result["cur"].get("mtime", "")),
             "- 요약: " + (" · ".join("%s %d" % (KIND_KO.get(k, k), v) for k, v in s.items() if v and k != "total") or "차이 없음")]
    lines += ["- 참고: " + n for n in result.get("notes", [])] + [""]
    sh = result["sheets"]
    for key, label in (("added", "시트 추가"), ("removed", "시트 삭제")):
        if sh[key]:
            lines.append("- %s: %s" % (label, ", ".join(sh[key])))
    if sh["renamed"]:
        lines.append("- 시트 이름 변경: " + ", ".join("%s → %s" % p for p in sh["renamed"]))
    if sh["moved"]:
        lines.append("- 시트 순서 변경")
    by = {}
    for c in result["changes"]:
        by.setdefault(c["sheet"], []).append(c)
    for name, cs in by.items():
        lines += ["", "## %s" % name, "", "| 종류 | 위치 | 기준 위치 | 내용 | 코드 |", "|---|---|---|---|---|"]
        for c in cs:
            k = c["kind"]
            where = c.get("at") or ", ".join(c.get("range", []))
            base_where = c.get("base_at") or ", ".join(c.get("base_range", []))
            if k in ("row_insert", "col_insert", "row_delete", "col_delete"):
                body = "%d개 · %s" % (c.get("count", 0), _fmt(", ".join(c.get("preview", [])), 80))
            elif k in ("cf", "drawing") or c.get("attr") in ("add", "remove"):
                body = "%s %s → %s" % (c.get("attr") or "", _fmt(c.get("before")), _fmt(c.get("after")))
            else:
                x, y = _pair(c.get("before"), c.get("after"))
                body = "%s%s → %s" % ((c["attr"] + " ") if c.get("attr") else "", x, y)
            code = c.get("code", {})
            ref = "<br>".join(["쓴 곳 " + x for x in code.get("set_at", [])] + ["문구 " + x for x in code.get("literal", [])])
            lines.append("| %s | %s | %s | %s | %s |" % (KIND_KO.get(k, k), _fmt(where, 40), _fmt(base_where, 40), body, ref))
    if result.get("names"):
        lines += ["", "## 정의된 이름", ""]
        for n in result["names"]:
            lines.append("- %s: %s → %s%s" % (n["name"], _fmt(n["before"]), _fmt(n["after"]),
                                               ("  (쓴 곳 %s)" % n["code"]["set_at"][0]) if n.get("code", {}).get("set_at") else ""))
    if result.get("vba"):
        lines += ["", "## VBA", ""]
        for v in result["vba"]:
            lines += ["### %s (%s)" % (v["module"], {"added": "추가", "removed": "삭제", "changed": "변경"}[v["kind"]]), "```vb"] + v["diff"][:60] + ["```"]
    return "\n".join(lines) + "\n"
