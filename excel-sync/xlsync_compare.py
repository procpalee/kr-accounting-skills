# -*- coding: utf-8 -*-
"""xlsync_compare - 스냅샷 두 개(기준 A -> 현재 B)의 차이. 같은 입력이면 항상 같은 결과·같은 순서.

1) 시트 짝: 이름이 같으면 짝, 없어진 이름과 새 이름은 값 겹침 60% 이상이면 이름 바뀜.
2) 행 맞춤: 행마다 값 묶음(열 위치 무시, 수식은 칸 참조 제거)으로 서명을 만들어 difflib 로 맞춘다 -> 삽입·삭제 행.
   열 맞춤: 짝지은 행에서 열 서명을 만들어 같은 방법(열 절반 이상이 바뀌면 열 구조 변경으로 보지 않음).
3) 기준 수식·이름·유효성 수식·병합·조건부 서식 범위는 행·열 이동과 시트 이름 변경을 엑셀 규칙대로 옮긴 뒤 비교한다.
4) 칸 비교: 값·수식 / 서식 속성(font.b·fill.fg·border.left·align.h·nf ...). 같은 (속성, 전, 후) 칸은 범위로 묶는다.
5) 이진 파트(printerSettings 등)는 해시로 같음/바뀜만 본다. 한쪽이 옛 스냅샷(기록 없음)이면 건너뛰고 notes 에 적는다.
"""
import difflib
import json

from xlsync_refs import AxisMap, a1, column_index_from_string, expand, get_column_letter, group, rc, strip_refs, translate_formula

WIDTH_TOL, HEIGHT_TOL = 0.06, 0.05      # 엑셀은 열너비를 픽셀 단위로 반올림해 저장한다(23.7109 -> 23.75)
KINDS = ("value", "style", "width", "height", "merge", "dv", "cf", "comment", "link", "checkbox", "drawing", "binary", "sheet", "names", "vba")
OLD_BIN = "한쪽 스냅샷이 옛 형식(이진 파트 기록 없음)이라 이진 파트(printerSettings 등)는 비교하지 않음 - 다시 snap 하면 비교"


def jk(v):
    return json.dumps(v, sort_keys=True, ensure_ascii=False)


def _norm(v):
    if isinstance(v, dict):
        return "f:" + str(strip_refs(v["f"])) if "f" in v else "d:" + jk(v)
    return "%s:%s" % (type(v).__name__[0], v)


def extent(sh):
    mr = mc = 0
    for addr in list(sh.get("cells", {})) + list(sh.get("dv", {})) + list(sh.get("checkbox", {})):
        r, c = rc(addr)
        mr, mc = max(mr, r), max(mc, c)
    for r in sh.get("rows", {}):
        mr = max(mr, int(r))
    for L in sh.get("cols", {}):
        mc = max(mc, column_index_from_string(L))
    for m in sh.get("merges", []):
        r, c = rc(m.split(":")[-1])
        mr, mc = max(mr, r), max(mc, c)
    return mr, mc


def _sim(x, y):
    if not x and not y:
        return 1.0
    sx, sy = set(x), set(y)
    return len(sx & sy) / float(len(sx | sy) or 1)


def _pair_block(a, b):
    n, m = len(a), len(b)
    if n * m > 90000:
        return [(i, i) for i in range(min(n, m))]
    dp = [[0.0] * (m + 1) for _ in range(n + 1)]
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            s = _sim(a[i - 1], b[j - 1])
            best = max(dp[i - 1][j], dp[i][j - 1])
            if s >= 0.5:
                best = max(best, dp[i - 1][j - 1] + s)
            dp[i][j] = best
    pairs, i, j = [], n, m
    while i > 0 and j > 0:
        s = _sim(a[i - 1], b[j - 1])
        if s >= 0.5 and abs(dp[i][j] - (dp[i - 1][j - 1] + s)) < 1e-9:
            pairs.append((i - 1, j - 1))
            i, j = i - 1, j - 1
        elif dp[i - 1][j] >= dp[i][j - 1]:
            i -= 1
        else:
            j -= 1
    return pairs[::-1]


def align(a, b):
    keep = {}
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1):
            for k in range(i2 - i1):
                keep[i1 + k + 1] = j1 + k + 1
        elif tag == "replace":
            for x, y in _pair_block(a[i1:i2], b[j1:j2]):
                keep[i1 + x + 1] = j1 + y + 1
    return keep


def _row_sigs(sh, n):
    rows = {}
    for addr, e in sh.get("cells", {}).items():
        if "v" in e:
            rows.setdefault(rc(addr)[0], []).append(_norm(e["v"]))
    return [tuple(sorted(rows.get(r, ()))) for r in range(1, n + 1)]


def _col_sigs(sh, rows, n):
    pos = {r: i for i, r in enumerate(rows)}
    cols = {}
    for addr, e in sh.get("cells", {}).items():
        if "v" in e:
            r, c = rc(addr)
            if r in pos:
                cols.setdefault(c, []).append((pos[r], _norm(e["v"])))
    return [tuple(sorted(cols.get(c, ()))) for c in range(1, n + 1)]


def axis_maps(X, Y):
    (nra, nca), (nrb, ncb) = extent(X), extent(Y)
    rk = align(_row_sigs(X, nra), _row_sigs(Y, nrb))
    pairs = sorted(rk.items())
    sa, sb = _col_sigs(X, [p[0] for p in pairs], nca), _col_sigs(Y, [p[1] for p in pairs], ncb)
    ck = align(sa, sb)
    same = sum(1 for i, j in ck.items() if sa[i - 1] and sa[i - 1] == sb[j - 1])
    if sum(1 for s in sa if s) and same < 0.5 * sum(1 for s in sa if s):
        ck = {i: i for i in range(1, min(nca, ncb) + 1)}
    return AxisMap(rk, nra), AxisMap(ck, nca), (nra, nca, nrb, ncb)


def match_sheets(A, B):
    sa, sb = {s["name"]: s for s in A["sheets"]}, {s["name"]: s for s in B["sheets"]}
    na, nb = [s["name"] for s in A["sheets"]], [s["name"] for s in B["sheets"]]
    pairs = [(n, n) for n in na if n in sb]
    ra, rb = [n for n in na if n not in sb], [n for n in nb if n not in sa]

    def vals(sh):
        return {jk(e["v"]) for e in sh.get("cells", {}).values() if "v" in e}
    cand = []
    for x in ra:
        vx = vals(sa[x])
        for y in rb:
            vy = vals(sb[y])
            s = len(vx & vy) / float(len(vx | vy) or 1)
            if s >= 0.6:
                cand.append((-s, x, y))
    renames = {}
    for _, x, y in sorted(cand):
        if x not in renames and y not in renames.values():
            renames[x] = y
            pairs.append((x, y))
    pairs.sort(key=lambda p: na.index(p[0]))
    return pairs, renames, [y for y in rb if y not in renames.values()], [x for x in ra if x not in renames]


def flat(style):
    out = {}
    for g, d in (style or {}).items():
        if isinstance(d, dict):
            for k, v in d.items():
                out["%s.%s" % (g, k)] = v
        else:
            out[g] = d
    return out


def same_value(a, b):
    if type(a) is bool or type(b) is bool:
        return type(a) is type(b) and a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= 1e-9 * max(1.0, abs(float(a)))
    return a == b


def _spans(nums):
    out, s = [], None
    for n in sorted(nums):
        if s is None or n != p + 1:
            if s is not None:
                out.append((s, p))
            s = n
        p = n
    if s is not None:
        out.append((s, p))
    return out


def _tr_range(rng, rm, cm):
    p, q = (rng.split(":") + [rng])[:2] if ":" in rng else (rng, rng)
    (r1, c1), (r2, c2) = rc(p), rc(q)
    r1, r2, c1, c2 = rm.start(r1), rm.end(r2), cm.start(c1), cm.end(c2)
    if None in (r1, r2, c1, c2) or r1 > r2 or c1 > c2:
        return None
    return a1(r1, c1) if (r1, c1) == (r2, c2) else "%s:%s" % (a1(r1, c1), a1(r2, c2))


def _tr_text(f, sheet, maps, renames):
    if not isinstance(f, str) or not f:
        return f
    return translate_formula(f if f.startswith("=") else "=" + f, sheet, maps, renames)[0 if f.startswith("=") else 1:]


def compare(A, B, ignore=()):
    ig = set(ignore)
    sa, sb = {s["name"]: s for s in A["sheets"]}, {s["name"]: s for s in B["sheets"]}
    pairs, renames, added, removed = match_sheets(A, B)
    res = {"base": A["meta"], "cur": B["meta"], "sheets": {"added": added, "removed": removed, "renamed": sorted(renames.items()),
           "moved": [p[1] for p in pairs] != [n for n in [s["name"] for s in B["sheets"]] if n in dict((y, x) for x, y in pairs)]},
           "changes": [], "names": [], "vba": [], "maps": {}, "notes": []}
    maps, dims = {}, {}
    for na, nb in pairs:
        if sa[na].get("chartsheet") or sb[nb].get("chartsheet"):
            continue
        rm, cm, d = axis_maps(sa[na], sb[nb])
        maps[na], dims[na] = (rm, cm), d
    for na, nb in pairs:
        if na in maps:
            _sheet(res, A, B, sa[na], sb[nb], maps, renames, dims[na], ig)
        elif "sheet" not in ig and sa[na].get("state") != sb[nb].get("state"):
            res["changes"].append({"sheet": nb, "base_sheet": na, "kind": "sheet", "attr": "state", "before": sa[na].get("state"), "after": sb[nb].get("state")})
    if "names" not in ig:
        na_, nb_ = A.get("names", {}), B.get("names", {})
        for n in sorted(set(na_) | set(nb_)):
            before = _tr_text(na_.get(n), None, maps, renames) if n in na_ else None
            if n not in nb_ or n not in na_ or before != nb_[n]:
                res["names"].append({"name": n, "before": na_.get(n), "before_moved": before, "after": nb_.get(n)})
    if "vba" not in ig:
        va, vb = A.get("vba") or {}, B.get("vba") or {}
        for m in sorted(set(va) | set(vb)):
            if va.get(m) != vb.get(m):
                lines = list(difflib.unified_diff((va.get(m) or "").split("\n"), (vb.get(m) or "").split("\n"), "기준", "현재", n=1, lineterm=""))
                res["vba"].append({"module": m, "kind": "added" if m not in va else "removed" if m not in vb else "changed", "diff": lines[:80]})
    cnt = {}
    for c in res["changes"]:
        cnt[c["kind"]] = cnt.get(c["kind"], 0) + 1
    res["summary"] = dict(cnt, names=len(res["names"]), vba=len(res["vba"]), sheets_added=len(added), sheets_removed=len(removed),
                          sheets_renamed=len(renames))
    res["summary"]["total"] = len(res["changes"]) + len(res["names"]) + len(res["vba"]) + len(added) + len(removed) + len(renames) + int(res["sheets"]["moved"])
    return res


def _sheet(res, A, B, X, Y, maps, renames, dims, ig):
    na, nb = X["name"], Y["name"]
    rm, cm = maps[na]
    nra, nca, nrb, ncb = dims
    ch = res["changes"]
    kept_r, kept_c = set(rm.keep.values()) if not rm.identity else set(range(1, nrb + 1)), set(cm.keep.values()) if not cm.identity else set(range(1, ncb + 1))
    inv_r = {v: k for k, v in rm.keep.items()}
    inv_c = {v: k for k, v in cm.keep.items()}
    res["maps"][nb] = {"rows": [] if rm.identity else sorted(rm.keep.items()), "cols": [] if cm.identity else sorted(cm.keep.items())}

    def base_of(cells):
        return group([(inv_r.get(r, r), inv_c.get(c, c)) for r, c in cells])

    def add(kind, cells, **kw):
        ch.append(dict({"sheet": nb, "base_sheet": na, "kind": kind, "range": group(cells), "base_range": base_of(cells), "n": len(cells)}, **kw))

    for tag, axis, nums, src in (("row_insert", "행", [r for r in range(1, nrb + 1) if r not in kept_r], Y),
                                 ("row_delete", "행", [r for r in range(1, nra + 1) if r not in rm.keep and not rm.identity], X),
                                 ("col_insert", "열", [c for c in range(1, ncb + 1) if c not in kept_c], Y),
                                 ("col_delete", "열", [c for c in range(1, nca + 1) if c not in cm.keep and not cm.identity], X)):
        for s, e in _spans(nums):
            is_row = tag.startswith("row")
            label = ("%d" % s if s == e else "%d:%d" % (s, e)) if is_row else (get_column_letter(s) if s == e else "%s:%s" % (get_column_letter(s), get_column_letter(e)))
            preview = []
            for addr, en in sorted(src.get("cells", {}).items(), key=lambda kv: rc(kv[0])):
                r, c = rc(addr)
                if "v" in en and ((is_row and s <= r <= e) or (not is_row and s <= c <= e)):
                    preview.append("%s=%s" % (addr, en["v"]["f"] if isinstance(en["v"], dict) and "f" in en["v"] else en["v"]))
            ch.append({"sheet": nb, "base_sheet": na, "kind": tag, "at": label, "count": e - s + 1, "preview": preview[:8]})

    CA, CB, STA, STB = X.get("cells", {}), Y.get("cells", {}), A["styles"], B["styles"]
    vals, styles, seen = {}, {}, set()
    for addr, ea in CA.items():
        r, c = rc(addr)
        r2, c2 = rm.one(r), cm.one(c)
        if r2 is None or c2 is None:
            continue
        b = a1(r2, c2)
        seen.add(b)
        eb = CB.get(b, {})
        _cell(ea, eb, X, Y, STA, STB, na, maps, renames, vals, styles, (r2, c2), ig)
    for b, eb in CB.items():
        if b in seen:
            continue
        r2, c2 = rc(b)
        if r2 in kept_r and c2 in kept_c:
            _cell({}, eb, X, Y, STA, STB, na, maps, renames, vals, styles, (r2, c2), ig)
    for (bv, av), cells in sorted(vals.items(), key=lambda kv: sorted(kv[1])[0]):
        add("value", cells, before=json.loads(bv), after=json.loads(av))
    for (attr, bv, av), cells in sorted(styles.items(), key=lambda kv: (sorted(kv[1])[0], kv[0][0])):
        add("style", cells, attr=attr, before=json.loads(bv), after=json.loads(av))

    if "width" not in ig:
        for axis, da, db, mp, conv, back, tol, key in (("col", X.get("cols", {}), Y.get("cols", {}), cm, column_index_from_string, get_column_letter, WIDTH_TOL, "w"),
                                                       ("row", X.get("rows", {}), Y.get("rows", {}), rm, int, str, HEIGHT_TOL, "h")):
            if axis == "row" and "height" in ig:
                continue
            done = set()
            for k, d in da.items():
                j = mp.one(conv(k))
                if j is None:
                    continue
                kb = back(j)
                done.add(kb)
                _dim(ch, nb, na, axis, kb, k, d, db.get(kb, {}), tol, key)
            for kb, d in db.items():
                if kb not in done:
                    i = (inv_c if axis == "col" else inv_r).get(conv(kb), conv(kb))
                    if (axis == "col" and conv(kb) in kept_c) or (axis == "row" and conv(kb) in kept_r):
                        _dim(ch, nb, na, axis, kb, back(i), {}, d, tol, key)

    if "merge" not in ig:
        ma = {t for t in (_tr_range(m, rm, cm) for m in X.get("merges", [])) if t}
        mb = set(Y.get("merges", []))
        for m in sorted(mb - ma):
            ch.append({"sheet": nb, "base_sheet": na, "kind": "merge", "attr": "add", "range": [m]})
        for m in sorted(ma - mb):
            ch.append({"sheet": nb, "base_sheet": na, "kind": "merge", "attr": "remove", "range": [m]})

    for kind, key in (("dv", "dv"), ("comment", "comments"), ("link", "links"), ("checkbox", "checkbox")):
        if kind in ig:
            continue
        da, db, got = X.get(key, {}), Y.get(key, {}), {}
        seen = set()

        def sig(v, side):
            if kind != "dv" or v is None:
                return v
            d = dict((A if side == "a" else B)["dvs"][v])
            if side == "a":
                for f in ("f1", "f2"):
                    if f in d:
                        d[f] = _tr_text(d[f], na, maps, renames)
            return d
        for addr, v in da.items():
            r, c = rc(addr)
            r2, c2 = rm.one(r), cm.one(c)
            if r2 is None or c2 is None:
                continue
            b = a1(r2, c2)
            seen.add(b)
            x, y = sig(v, "a"), sig(db.get(b), "b")
            if jk(x) != jk(y):
                got.setdefault((jk(x), jk(y)), []).append((r2, c2))
        for b, v in db.items():
            r2, c2 = rc(b)
            if b not in seen and r2 in kept_r and c2 in kept_c:
                got.setdefault((jk(None), jk(sig(v, "b"))), []).append((r2, c2))
        for (x, y), cells in sorted(got.items(), key=lambda kv: sorted(kv[1])[0]):
            add(kind, cells, before=json.loads(x), after=json.loads(y))

    if "cf" not in ig:
        def cf_key(sq, rule, moved):
            if moved:
                sq = " ".join(t for t in (_tr_range(p, rm, cm) for p in sq.split()) if t)
                rule = dict(rule, f=[_tr_text(f, na, maps, renames) for f in rule.get("f", [])])
            return jk([sq, rule])
        ka = sorted(cf_key(sq, rule, True) for sq, rule in X.get("cf", []))
        kb = sorted(cf_key(sq, rule, False) for sq, rule in Y.get("cf", []))
        for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, ka, kb, autojunk=False).get_opcodes():
            if tag in ("delete", "replace"):
                for k in ka[i1:i2]:
                    ch.append({"sheet": nb, "base_sheet": na, "kind": "cf", "attr": "remove", "before": json.loads(k)})
            if tag in ("insert", "replace"):
                for k in kb[j1:j2]:
                    ch.append({"sheet": nb, "base_sheet": na, "kind": "cf", "attr": "add", "after": json.loads(k)})

    if "drawing" not in ig and X.get("drawing", {}) != Y.get("drawing", {}):
        ch.append({"sheet": nb, "base_sheet": na, "kind": "drawing", "before": X.get("drawing", {}), "after": Y.get("drawing", {})})
    if "binary" not in ig:
        if "bin" in X and "bin" in Y:
            for k in sorted(set(X["bin"]) | set(Y["bin"])):
                if X["bin"].get(k) != Y["bin"].get(k):
                    ch.append({"sheet": nb, "base_sheet": na, "kind": "binary", "at": k, "before": X["bin"].get(k), "after": Y["bin"].get(k)})
        elif OLD_BIN not in res["notes"]:
            res["notes"].append(OLD_BIN)
    if "sheet" not in ig:
        fa = X.get("freeze")
        fa = _tr_range(fa, rm, cm) if fa else fa
        for attr, x, y in (("state", X.get("state"), Y.get("state")), ("tab", X.get("tab"), Y.get("tab")),
                           ("freeze", fa, Y.get("freeze")), ("grid", X.get("grid"), Y.get("grid")),
                           ("print", X.get("print"), Y.get("print"))):
            if jk(x) != jk(y):
                ch.append({"sheet": nb, "base_sheet": na, "kind": "sheet", "attr": attr, "before": x, "after": y})


def _cell(ea, eb, X, Y, STA, STB, na, maps, renames, vals, styles, at, ig):
    if "value" not in ig:
        va, vb = ea.get("v"), eb.get("v")
        if isinstance(va, dict) and "f" in va:
            va = dict(va, f=translate_formula(va["f"], na, maps, renames))
        if not same_value(va, vb):
            vals.setdefault((jk(va), jk(vb)), []).append(at)
    if "style" not in ig:
        fa = flat(STA.get(ea.get("s", X["default"]), {}))
        fb = flat(STB.get(eb.get("s", Y["default"]), {}))
        if fa != fb:
            for k in sorted(set(fa) | set(fb)):
                if jk(fa.get(k)) != jk(fb.get(k)):
                    styles.setdefault((k, jk(fa.get(k)), jk(fb.get(k))), []).append(at)


def _dim(ch, nb, na, axis, kb, ka, da, db, tol, key):
    for attr in (key, "hidden"):
        x, y = da.get(attr), db.get(attr)
        if attr == key and x is not None and y is not None and abs(x - y) <= tol:
            continue
        if x != y:
            ch.append({"sheet": nb, "base_sheet": na, "kind": "width" if axis == "col" else "height",
                       "attr": "width" if attr == "w" else "height" if attr == "h" else "hidden",
                       "at": kb, "base_at": ka, "before": x, "after": y})
