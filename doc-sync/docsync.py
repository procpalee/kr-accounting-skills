# -*- coding: utf-8 -*-
"""docsync - Word·PPT 를 사람이 직접 고치고 저장하면, 기준본과 결정론적으로 비교해 바뀐 점과 고칠 코드 위치를 찾는다.

  python -X utf8 docsync.py snap   <문서> [--label 설명]
  python -X utf8 docsync.py status <문서>
  python -X utf8 docsync.py diff   <문서> [--base 기준(.json/.docx/.pptx)] [--code 코드폴더 ...] [--ignore format,geom,...]
  python -X utf8 docsync.py verify <기대 문서(사람이 고친 것)> <실제 문서(코드로 다시 만든 것)> [--ignore ...]

결과는 문서 옆 _docsync/<문서 이름>/ 에만 남는다(base.json·원본 사본·diff_시각.md/.json·history/).
Word 는 문단·표 칸, PPT 는 슬라이드·도형 단위로 본다. 표준 라이브러리만 쓴다.
"""
import argparse
import datetime
import difflib
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
KINDS = ("text", "format", "insert", "delete", "move", "table", "page", "geom", "name", "image", "chart", "notes")


def sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def store_dir(path):
    d = os.path.join(os.path.dirname(path), "_docsync", os.path.basename(path))
    os.makedirs(d, exist_ok=True)
    return d


# ── Word 읽기 ─────────────────────────────────────────────────────────────
def _val(el, name=W + "val"):
    return el.get(name) if el is not None else None


def _on(rpr, tag):
    el = rpr.find(W + tag)
    return el is not None and el.get(W + "val") not in ("0", "false", "none")


_STYLE_NAMES = {}      # 읽는 문서의 스타일 ID -> 이름


def _style(style_id):
    return _STYLE_NAMES.get(style_id, style_id) if style_id else style_id


def w_fmt(rpr):
    """글자 서식을 비교용 문자열로. 언어·교정·복합 글꼴 표시처럼 눈에 안 보이는 것은 뺀다."""
    if rpr is None:
        return ""
    out = [t for t in ("b", "i", "strike") if _on(rpr, t)]
    for tag in ("u", "color", "sz", "highlight", "vertAlign", "rStyle"):
        v = _val(rpr.find(W + tag))
        if v and v != "none":
            out.append("%s=%s" % (tag, _style(v) if tag == "rStyle" else v))
    f = rpr.find(W + "rFonts")
    if f is not None:
        v = f.get(W + "eastAsia") or f.get(W + "ascii")
        if v:
            out.append("font=" + v)
    sh = rpr.find(W + "shd")
    if sh is not None and sh.get(W + "fill") not in (None, "auto"):
        out.append("shd=" + sh.get(W + "fill"))
    return ";".join(sorted(out))


def w_ppr(p):
    ppr = p.find(W + "pPr")
    out = {}
    if ppr is None:
        return out
    v = _val(ppr.find(W + "pStyle"))
    if v:
        out["style"] = _style(v)
    v = _val(ppr.find(W + "jc"))
    if v:
        out["align"] = v
    for tag, keys in (("ind", ("left", "right", "firstLine", "hanging", "start", "end")), ("spacing", ("before", "after", "line"))):
        el = ppr.find(W + tag)
        if el is not None:
            for k in keys:
                if el.get(W + k) is not None:
                    out["%s.%s" % (tag, k)] = el.get(W + k)
    num = ppr.find(W + "numPr")
    if num is not None:
        out["num"] = "%s/%s" % (_val(num.find(W + "numId")), _val(num.find(W + "ilvl")))
    return out


def _iter_visible(el):
    """대체 표시(mc:Fallback)는 같은 내용의 복사본이라 건너뛴다."""
    for k in el:
        if k.tag == MC_FALLBACK:
            continue
        yield k
        yield from _iter_visible(k)


def _spans(p):
    spans, slots, images = [], [], 0

    def walk(node, slot):
        nonlocal images
        for k in node:
            t = k.tag
            if t == W + "r":
                fmt = w_fmt(k.find(W + "rPr"))
                buf = []
                for x in k:
                    if x.tag == W + "t":
                        buf.append(x.text or "")
                    elif x.tag in (W + "tab", W + "ptab"):
                        buf.append("\t")
                    elif x.tag in (W + "br", W + "cr"):
                        buf.append("\n")
                    elif x.tag == W + "noBreakHyphen":
                        buf.append("-")
                    elif x.tag in (W + "drawing", W + "pict"):
                        images += 1
                    elif x.tag == "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent":
                        images += 1
                if buf:
                    text = "".join(buf)
                    if spans and spans[-1][1] == fmt:
                        spans[-1][0] += text
                    else:
                        spans.append([text, fmt])
            elif t == W + "sdt":
                name = _val(k.find(W + "sdtPr").find(W + "tag")) or _val(k.find(W + "sdtPr").find(W + "alias"))
                if name:
                    slots.append(name)
                content = k.find(W + "sdtContent")
                if content is not None:
                    walk(content, name or slot)
            elif t == W + "bookmarkStart":
                n = k.get(W + "name") or ""
                if n and not n.startswith("_"):
                    slots.append(n)
            elif t in (W + "hyperlink", W + "ins", W + "smartTag", W + "fldSimple"):
                walk(k, slot)
    walk(p, None)
    return spans, slots, images


def w_para(p, slot):
    spans, slots, images = _spans(p)
    d = {"k": "p", "text": "".join(s[0] for s in spans), "spans": spans, "ppr": w_ppr(p)}
    if images:
        d["img"] = images
    name = slot or (slots[0] if slots else None)
    if name:
        d["slot"] = name
    return d


def w_table(tbl, slot):
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            pr = tc.find(W + "tcPr")
            c = {"blocks": list(w_blocks(tc, slot))}
            if pr is not None:
                span = _val(pr.find(W + "gridSpan"))
                if span and span != "1":
                    c["span"] = span
                sh = pr.find(W + "shd")
                if sh is not None and sh.get(W + "fill") not in (None, "auto"):
                    c["shade"] = sh.get(W + "fill")
                if pr.find(W + "vMerge") is not None:
                    c["vmerge"] = _val(pr.find(W + "vMerge")) or "continue"
            cells.append(c)
        rows.append(cells)
    d = {"k": "tbl", "rows": rows}
    st = tbl.find(W + "tblPr")
    if st is not None and _val(st.find(W + "tblStyle")):
        d["style"] = _style(_val(st.find(W + "tblStyle")))
    if slot:
        d["slot"] = slot
    return d


def w_blocks(container, slot=None):
    for el in container:
        if el.tag == W + "p":
            yield w_para(el, slot)
            for tx in (x for x in _iter_visible(el) if x.tag == W + "txbxContent"):
                yield from w_blocks(tx, slot)
        elif el.tag == W + "tbl":
            yield w_table(el, slot)
        elif el.tag == W + "sdt":
            pr = el.find(W + "sdtPr")
            name = (_val(pr.find(W + "tag")) or _val(pr.find(W + "alias"))) if pr is not None else None
            content = el.find(W + "sdtContent")
            if content is not None:
                yield from w_blocks(content, name or slot)
        elif el.tag in (W + "ins", W + "customXml"):
            yield from w_blocks(el, slot)


def w_page(body):
    out = []
    for s in body.iter(W + "sectPr"):
        d = {}
        sz, mar = s.find(W + "pgSz"), s.find(W + "pgMar")
        if sz is not None:
            d.update(w=sz.get(W + "w"), h=sz.get(W + "h"), orient=sz.get(W + "orient") or "portrait")
        if mar is not None:
            d.update({k: mar.get(W + k) for k in ("top", "bottom", "left", "right")})
        out.append(d)
    return out


def snap_docx(path):
    parts, page = {}, []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        _STYLE_NAMES.clear()
        if "word/styles.xml" in names:
            for st in ET.fromstring(z.read("word/styles.xml")).findall(W + "style"):
                n = st.find(W + "name")
                if n is not None:
                    _STYLE_NAMES[st.get(W + "styleId")] = n.get(W + "val").lower()
        order = [("word/document.xml", "본문")]
        for prefix, label in (("word/header", "머리글"), ("word/footer", "바닥글")):
            found = sorted(n for n in names if re.fullmatch(re.escape(prefix) + r"\d*\.xml", n))
            order += [(n, label + (re.search(r"(\d*)\.xml$", n).group(1) or "")) for n in found]
        order += [("word/footnotes.xml", "각주"), ("word/endnotes.xml", "미주")]
        for name, label in order:
            if name not in names:
                continue
            root = ET.fromstring(z.read(name))
            if name == "word/document.xml":
                body = root.find(W + "body")
                parts[label] = list(w_blocks(body))
                page = w_page(body)
            elif label in ("각주", "미주"):
                blocks = []
                for note in root:
                    if note.get(W + "type") in (None, "normal"):
                        blocks += list(w_blocks(note))
                parts[label] = blocks
            else:
                parts[label] = list(w_blocks(root))
    return {"parts": parts, "page": page}


# ── PPT 읽기 ──────────────────────────────────────────────────────────────
def _rels(z, part):
    d, f = os.path.split(part)
    rp = "%s/_rels/%s.rels" % (d, f)
    if rp not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(z.read(rp)):
        tgt = r.get("Target") or ""
        if r.get("TargetMode") == "External":
            continue
        full = tgt.lstrip("/") if tgt.startswith("/") else os.path.normpath(os.path.join(d, tgt)).replace("\\", "/")
        out[r.get("Id")] = (full, (r.get("Type") or "").rsplit("/", 1)[-1])
    return out


def a_fmt(rpr):
    if rpr is None:
        return ""
    out = []
    for k in ("b", "i", "strike", "u"):
        v = rpr.get(k)
        if v not in (None, "0", "none", "noStrike"):
            out.append("%s=%s" % (k, v))
    if rpr.get("sz"):
        out.append("sz=" + rpr.get("sz"))
    fill = rpr.find(A + "solidFill")
    if fill is not None and len(fill):
        out.append("color=" + (fill[0].get("val") or ""))
    for tag in ("ea", "latin"):
        f = rpr.find(A + tag)
        if f is not None and f.get("typeface"):
            out.append("font=" + f.get("typeface"))
            break
    return ";".join(sorted(out))


def a_para(p):
    spans = []
    for k in p:
        if k.tag in (A + "r", A + "fld"):
            text, fmt = (k.findtext(A + "t") or ""), a_fmt(k.find(A + "rPr"))
        elif k.tag == A + "br":
            text, fmt = "\n", (spans[-1][1] if spans else "")
        else:
            continue
        if spans and spans[-1][1] == fmt:
            spans[-1][0] += text
        else:
            spans.append([text, fmt])
    ppr, out = p.find(A + "pPr"), {}
    if ppr is not None:
        for k in ("algn", "lvl", "marL", "indent"):
            if ppr.get(k) is not None:
                out[k] = ppr.get(k)
    return {"k": "p", "text": "".join(s[0] for s in spans), "spans": spans, "ppr": out}


def a_body(body):
    return [a_para(p) for p in body.findall(A + "p")] if body is not None else []


def p_shape(z, rels, el, prefix=""):
    tag = el.tag.rsplit("}", 1)[-1]
    nv = next((x for x in el.iter(P + "cNvPr")), None)
    name = nv.get("name") if nv is not None else ""
    d = {"id": nv.get("id") if nv is not None else "", "name": prefix + name, "type": tag}
    xf = None
    for k in el:
        if k.tag == P + "spPr" or k.tag == P + "grpSpPr":
            xf = k.find(A + "xfrm")
        elif k.tag == P + "xfrm":
            xf = k
    if xf is not None and xf.find(A + "off") is not None:
        off, ext = xf.find(A + "off"), xf.find(A + "ext")
        d["geom"] = [int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))]
    sp = el.find(P + "spPr")
    if sp is not None and sp.find(A + "solidFill") is not None and len(sp.find(A + "solidFill")):
        d["fill"] = sp.find(A + "solidFill")[0].get("val")
    tx = el.find(P + "txBody")
    if tx is not None:
        d["paras"] = a_body(tx)
        bp = tx.find(A + "bodyPr")
        if bp is not None and bp.find(A + "spAutoFit") is not None:
            d["autofit"] = True                 # 글에 맞춰 PowerPoint 가 크기를 다시 정한다
    tbl = next((x for x in el.iter(A + "tbl")), None) if tag == "graphicFrame" else None
    if tbl is not None:
        d["rows"] = [[{"blocks": a_body(tc.find(A + "txBody")), **({"span": tc.get("gridSpan")} if tc.get("gridSpan") else {})}
                      for tc in tr.findall(A + "tc")] for tr in tbl.findall(A + "tr")]
    if tag == "pic":
        blip = next((x for x in el.iter(A + "blip")), None)
        rid = blip.get(R + "embed") if blip is not None else None
        if rid in rels:
            d["image"] = hashlib.sha1(z.read(rels[rid][0])).hexdigest()
    if tag == "graphicFrame":
        ch = next((x for x in el.iter(C + "chart")), None)
        if ch is not None and ch.get(R + "id") in rels:
            root = ET.fromstring(z.read(rels[ch.get(R + "id")][0]))
            d["chart"] = [v.text for v in root.iter(C + "v")]
    return d


def p_shapes(z, rels, tree, prefix=""):
    out = []
    for el in tree:
        if el.tag in (P + "sp", P + "pic", P + "graphicFrame", P + "cxnSp"):
            out.append(p_shape(z, rels, el, prefix))
        elif el.tag == P + "grpSp":
            g = p_shape(z, rels, el, prefix)
            out.append(g)
            out += p_shapes(z, rels, el, g["name"] + "/")
    return out


def snap_pptx(path):
    slides = []
    with zipfile.ZipFile(path) as z:
        pres = ET.fromstring(z.read("ppt/presentation.xml"))
        prels = _rels(z, "ppt/presentation.xml")
        lst = pres.find(P + "sldIdLst")
        for sid in (lst if lst is not None else []):
            part = prels[sid.get(R + "id")][0]
            rels = _rels(z, part)
            root = ET.fromstring(z.read(part))
            tree = root.find(P + "cSld").find(P + "spTree")
            s = {"sid": sid.get("id"), "shapes": p_shapes(z, rels, tree)}
            for tgt, typ in rels.values():
                if typ == "notesSlide":
                    nroot = ET.fromstring(z.read(tgt))
                    texts = []
                    for sp in nroot.iter(P + "sp"):
                        ph = next((x for x in sp.iter(P + "ph")), None)
                        if ph is not None and ph.get("type") == "body":
                            texts += [p["text"] for p in a_body(sp.find(P + "txBody"))]
                    s["notes"] = "\n".join(texts)
            slides.append(s)
    return {"slides": slides}


def snapshot(path, label=""):
    low = path.lower()
    if low.endswith(".docx"):
        s = snap_docx(path)
    elif low.endswith(".pptx"):
        s = snap_pptx(path)
    else:
        raise SystemExit("docx·pptx 만 다룹니다(엑셀은 excel-sync): " + path)
    s["meta"] = {"docsync": 1, "file": os.path.abspath(path), "sha1": sha1(path), "label": label,
                 "made": datetime.datetime.now().isoformat(timespec="seconds")}
    return {"meta": s.pop("meta"), **s}


def read_snap(path):
    if path.lower().endswith(".json"):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return snapshot(path)


# ── 비교 ──────────────────────────────────────────────────────────────────
def _block_text(b):
    if b["k"] == "p":
        return b["text"]
    return "\n".join("|".join(_cell_text(c) for c in r) for r in b["rows"])


def _cell_text(c):
    return "\n".join(_block_text(b) for b in c["blocks"])


def _ratio(a, b):
    if a == b:
        return 1.0
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


def align(xs, ys, floor=0.4):
    """두 목록을 순서대로 맞춘다. 반환: (짝 [(i, j)], 빠진 i, 새 j). 글이 일부 바뀐 항목은 유사도로 짝짓는다."""
    pairs, dels, ins = [], [], []
    sm = difflib.SequenceMatcher(None, xs, ys, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            pairs += list(zip(range(i1, i2), range(j1, j2)))
        elif tag == "delete":
            dels += range(i1, i2)
        elif tag == "insert":
            ins += range(j1, j2)
        else:
            j = j1
            for i in range(i1, i2):
                best, best_r = None, floor
                for k in range(j, j2):
                    r = _ratio(xs[i], ys[k])
                    if r > best_r:
                        best, best_r = k, r
                if best is None:
                    dels.append(i)
                else:
                    ins += range(j, best)
                    pairs.append((i, best))
                    j = best + 1
            ins += range(j, j2)
    return pairs, dels, ins


class Diff:
    def __init__(self, ignore=()):
        self.changes, self.ignore = [], set(ignore)

    def add(self, kind, at, before=None, after=None, base_at=None, slot=None, detail=None):
        if kind in self.ignore:
            return
        c = {"kind": kind, "at": at, "before": before, "after": after}
        if base_at and base_at != at:
            c["base_at"] = base_at
        if slot:
            c["slot"] = slot
        if detail:
            c["detail"] = detail
        self.changes.append(c)


def _chars(spans):
    return [f for t, f in spans for _ in t]


def _fmt_changes(a, b):
    """같은 글에서 서식이 달라진 구간. 반환: [(글, 전, 후)]."""
    ca, cb, text = _chars(a["spans"]), _chars(b["spans"]), b["text"]
    out, i = [], 0
    while i < len(ca):
        if ca[i] == cb[i]:
            i += 1
            continue
        j = i
        while j < len(ca) and ca[j] == ca[i] and cb[j] == cb[i]:
            j += 1
        out.append((text[i:j], ca[i], cb[i]))
        i = j
    return out


def _cmp_para(d, a, b, at, base_at):
    slot = b.get("slot") or a.get("slot")
    if a["text"] != b["text"]:
        d.add("text", at, a["text"], b["text"], base_at, slot)
    else:
        for seg, fa, fb in _fmt_changes(a, b):
            d.add("format", at, fa or "(기본)", fb or "(기본)", base_at, slot, "글자: " + seg[:30])
    if a.get("ppr") != b.get("ppr"):
        keys = sorted(k for k in set(a.get("ppr", {})) | set(b.get("ppr", {})) if a.get("ppr", {}).get(k) != b.get("ppr", {}).get(k))
        for k in keys:
            d.add("format", at, a.get("ppr", {}).get(k), b.get("ppr", {}).get(k), base_at, slot, "문단 " + k)
    if a.get("img") != b.get("img"):
        d.add("image", at, a.get("img", 0), b.get("img", 0), base_at, slot, "그림 수")


def _cmp_cell(d, a, b, at, base_at):
    if a.get("shade") != b.get("shade"):
        d.add("format", at, a.get("shade"), b.get("shade"), base_at, None, "칸 배경")
    if a.get("span") != b.get("span") or a.get("vmerge") != b.get("vmerge"):
        d.add("table", at, "%s/%s" % (a.get("span"), a.get("vmerge")), "%s/%s" % (b.get("span"), b.get("vmerge")), base_at, None, "병합")
    xs, ys = a["blocks"], b["blocks"]
    if [x["k"] for x in xs] == [y["k"] for y in ys] and len(xs) == 1:
        if xs[0]["k"] == "p":
            _cmp_para(d, xs[0], ys[0], at, base_at)
        return
    _cmp_blocks(d, xs, ys, at + " ", base_at + " " if base_at else None)


def _cmp_rows(d, ra, rb, label, base_label, slot=None):
    keys_a = ["|".join(_cell_text(c) for c in r) for r in ra]
    keys_b = ["|".join(_cell_text(c) for c in r) for r in rb]
    pairs, dels, ins = align(keys_a, keys_b)
    for i in dels:
        d.add("table", "%s %d행" % (label, i + 1), keys_a[i], None, slot=slot, detail="행 삭제")
    for j in ins:
        d.add("table", "%s %d행" % (label, j + 1), None, keys_b[j], slot=slot, detail="행 추가")
    for i, j in pairs:
        if len(ra[i]) != len(rb[j]):
            d.add("table", "%s %d행" % (label, j + 1), len(ra[i]), len(rb[j]), slot=slot, detail="열 수")
            continue
        for c, (x, y) in enumerate(zip(ra[i], rb[j]), 1):
            _cmp_cell(d, x, y, "%s %d행 %d열" % (label, j + 1, c), "%s %d행 %d열" % (base_label, i + 1, c))


def _cmp_blocks(d, xs, ys, part, base_part=None):
    base_part = base_part or part
    pairs, dels, ins = align([_block_text(x) for x in xs], [_block_text(y) for y in ys])
    tbl_no = lambda seq, n: sum(1 for b in seq[:n + 1] if b["k"] == "tbl")
    para_no = lambda seq, n: sum(1 for b in seq[:n + 1] if b["k"] == "p")

    def loc(seq, n, p):
        return "%s 표 %d" % (p, tbl_no(seq, n)) if seq[n]["k"] == "tbl" else "%s 문단 %d" % (p, para_no(seq, n))
    for i in dels:
        d.add("delete", loc(xs, i, base_part), _block_text(xs[i]), None, slot=xs[i].get("slot"))
    for j in ins:
        d.add("insert", loc(ys, j, part), None, _block_text(ys[j]), slot=ys[j].get("slot"))
    for i, j in pairs:
        a, b = xs[i], ys[j]
        if a["k"] != b["k"]:
            d.add("delete", loc(xs, i, base_part), _block_text(a), None)
            d.add("insert", loc(ys, j, part), None, _block_text(b))
        elif a["k"] == "p":
            _cmp_para(d, a, b, loc(ys, j, part), loc(xs, i, base_part))
        else:
            if a.get("style") != b.get("style"):
                d.add("format", loc(ys, j, part), a.get("style"), b.get("style"), loc(xs, i, base_part), None, "표 스타일")
            _cmp_rows(d, a["rows"], b["rows"], loc(ys, j, part), loc(xs, i, base_part), b.get("slot"))


def _shape_text(s):
    if "paras" in s:
        return "\n".join(p["text"] for p in s["paras"])
    if "rows" in s:
        return "\n".join("|".join(_cell_text(c) for c in r) for r in s["rows"])
    return ""


def _match_shapes(xs, ys):
    pairs, used_x, used_y = [], set(), set()

    def link(cond):
        for i, a in enumerate(xs):
            if i in used_x:
                continue
            for j, b in enumerate(ys):
                if j not in used_y and cond(a, b):
                    pairs.append((i, j))
                    used_x.add(i)
                    used_y.add(j)
                    break
    link(lambda a, b: a["id"] == b["id"] and a["name"] == b["name"])
    link(lambda a, b: a["name"] == b["name"])
    link(lambda a, b: a["type"] == b["type"] and _shape_text(a) and _shape_text(a) == _shape_text(b))
    link(lambda a, b: a["id"] == b["id"] and a["type"] == b["type"])
    return pairs, [i for i in range(len(xs)) if i not in used_x], [j for j in range(len(ys)) if j not in used_y]


def _cmp_shape(d, a, b, at, base_at):
    slot = b["name"]
    if a["name"] != b["name"]:
        d.add("name", at, a["name"], b["name"], base_at)
    ga, gb = a.get("geom"), b.get("geom")
    if ga and gb and (a.get("autofit") or b.get("autofit")):
        ga, gb = ga[:2], gb[:2]
    if ga != gb and not (ga and gb and all(abs(x - y) <= 1270 for x, y in zip(ga, gb))):
        d.add("geom", at, ga, gb, base_at, slot, "위치·크기")
    if a.get("fill") != b.get("fill"):
        d.add("format", at, a.get("fill"), b.get("fill"), base_at, slot, "채우기")
    if "paras" in a and "paras" in b:
        ta, tb = _shape_text(a), _shape_text(b)
        if ta != tb and len(a["paras"]) != len(b["paras"]):
            d.add("text", at, ta, tb, base_at, slot)
        else:
            for n, (x, y) in enumerate(zip(a["paras"], b["paras"]), 1):
                x2, y2 = dict(x), dict(y)
                x2["slot"] = y2["slot"] = slot
                label = at if len(b["paras"]) == 1 else "%s 문단 %d" % (at, n)
                _cmp_para(d, x2, y2, label, base_at if len(b["paras"]) == 1 else None)
    if "rows" in a and "rows" in b:
        _cmp_rows(d, a["rows"], b["rows"], at, base_at or at, slot)
    if a.get("image") != b.get("image"):
        d.add("image", at, a.get("image", "")[:8], b.get("image", "")[:8], base_at, slot, "그림 교체")
    if a.get("chart") != b.get("chart"):
        d.add("chart", at, a.get("chart"), b.get("chart"), base_at, slot, "차트 데이터")


def _slide_key(s):
    return "\n".join(_shape_text(x) for x in s["shapes"])


def _cmp_slides(d, xs, ys):
    pairs, dels, ins = align([_slide_key(s) for s in xs], [_slide_key(s) for s in ys], floor=0.3)
    moved = {}
    for i in list(dels):
        for j in list(ins):
            if _slide_key(xs[i]) == _slide_key(ys[j]):
                moved[i] = j
                dels.remove(i)
                ins.remove(j)
                break
    for i, j in moved.items():
        d.add("move", "슬라이드 %d" % (j + 1), i + 1, j + 1, "슬라이드 %d" % (i + 1), detail="순서 이동")
    for i in dels:
        d.add("delete", "슬라이드 %d" % (i + 1), _slide_key(xs[i])[:200], None, detail="슬라이드 삭제")
    for j in ins:
        d.add("insert", "슬라이드 %d" % (j + 1), None, _slide_key(ys[j])[:200], detail="슬라이드 추가")
    for i, j in pairs + list(moved.items()):
        a, b = xs[i], ys[j]
        sp, sd, si = _match_shapes(a["shapes"], b["shapes"])
        at_b, at_a = "슬라이드 %d" % (j + 1), "슬라이드 %d" % (i + 1)
        for n in sd:
            d.add("delete", "%s %s" % (at_a, a["shapes"][n]["name"]), _shape_text(a["shapes"][n]) or a["shapes"][n]["type"], None, detail="도형 삭제")
        for n in si:
            d.add("insert", "%s %s" % (at_b, b["shapes"][n]["name"]), None, _shape_text(b["shapes"][n]) or b["shapes"][n]["type"], detail="도형 추가")
        for x, y in sp:
            _cmp_shape(d, a["shapes"][x], b["shapes"][y], "%s %s" % (at_b, b["shapes"][y]["name"]), "%s %s" % (at_a, a["shapes"][x]["name"]))
        if a.get("notes", "") != b.get("notes", ""):
            d.add("notes", at_b, a.get("notes", ""), b.get("notes", ""), at_a)


def compare(A_, B_, ignore=()):
    d = Diff(ignore)
    if ("slides" in A_) != ("slides" in B_):
        raise SystemExit("문서 종류가 다릅니다")
    if "slides" in A_:
        _cmp_slides(d, A_["slides"], B_["slides"])
    else:
        for name in sorted(set(A_["parts"]) | set(B_["parts"]), key=lambda n: (n != "본문", n)):
            _cmp_blocks(d, A_["parts"].get(name, []), B_["parts"].get(name, []), name)
        if A_.get("page") != B_.get("page"):
            d.add("page", "쪽 설정", A_.get("page"), B_.get("page"))
    summary = {"total": len(d.changes)}
    for c in d.changes:
        summary[c["kind"]] = summary.get(c["kind"], 0) + 1
    return {"changes": d.changes, "summary": summary}


# ── 코드 위치·보고 ────────────────────────────────────────────────────────
def _code_tools():
    sys.path.insert(0, os.path.join(os.path.expanduser("~"), ".claude", "skills", "excel-sync"))
    from xlsync_report import code_index, find
    return code_index, find


def _candidates(before, after):
    """코드에 적혀 있을 만한 글 조각: 전체, 달라진 부분, 숫자로 끊은 가장 긴 조각."""
    if not isinstance(before, str):
        return []
    out = [before]
    if isinstance(after, str):
        i = 0
        while i < min(len(before), len(after)) and before[i] == after[i]:
            i += 1
        j = 0
        while j < min(len(before), len(after)) - i and before[-1 - j] == after[-1 - j]:
            j += 1
        frag = before[i:len(before) - j]
        if len(frag) >= 2:
            out.append(frag)
    chunks = sorted(re.split(r"[\d,.]+", before), key=len, reverse=True)
    if chunks and len(chunks[0].strip()) >= 4:
        out.append(chunks[0].strip())
    return out


def attach(res, roots, base):
    if not roots:
        return
    code_index, find = _code_tools()
    idx = code_index(roots)
    for c in res["changes"]:
        hits = []
        for cand in _candidates(c.get("before"), c.get("after")):
            hits = find(idx, cand, base)
            if hits:
                break
        if hits:
            c["code"] = hits


HOW = {"text": "글", "format": "서식", "insert": "추가", "delete": "삭제", "move": "이동", "table": "표 구조",
       "page": "쪽 설정", "geom": "위치·크기", "name": "도형 이름", "image": "그림", "chart": "차트", "notes": "발표 메모"}


def _how(c):
    if c.get("code"):
        return "코드 수정"
    if c.get("slot") and c["kind"] in ("text", "table", "image", "chart"):
        return "칸 값 확인"
    return "양식에 유지"


def _short(v, n=70):
    s = "" if v is None else (v if isinstance(v, str) else json.dumps(v, ensure_ascii=False))
    s = s.replace("\n", " / ").replace("|", "¦")
    return s if len(s) <= n else s[:n - 1] + "…"


def markdown(res, title="문서 변경 사항"):
    s = res["summary"]
    lines = ["# %s" % title, "", "총 %d건 (%s)" % (s["total"], ", ".join("%s %d" % (HOW[k], v) for k, v in s.items() if k != "total") or "없음"), ""]
    if not res["changes"]:
        return "\n".join(lines)
    lines += ["| 종류 | 위치 | 전 | 후 | 칸 | 반영 | 코드 |", "|---|---|---|---|---|---|---|"]
    for c in res["changes"]:
        kind = HOW[c["kind"]] + (" (%s)" % c["detail"] if c.get("detail") else "")
        lines.append("| %s | %s | %s | %s | %s | %s | %s |" % (kind, c["at"], _short(c.get("before")), _short(c.get("after")),
                                                               c.get("slot") or "", _how(c), "<br>".join(c.get("code", []))))
    return "\n".join(lines) + "\n"


# ── 명령 ──────────────────────────────────────────────────────────────────
def _ignore(s):
    ig = [x.strip() for x in (s or "").split(",") if x.strip()]
    bad = [x for x in ig if x not in KINDS]
    if bad:
        raise SystemExit("--ignore 에 모르는 종류: %s (가능: %s)" % (bad, ", ".join(KINDS)))
    return ig


def cmd_snap(a):
    path = os.path.abspath(a.file)
    d = store_dir(path)
    base = os.path.join(d, "base.json")
    copy = "base" + os.path.splitext(path)[1].lower()
    tmp_json, tmp_copy = os.path.join(d, "_tmp_base.json"), os.path.join(d, "_tmp_" + copy)
    try:
        s = snapshot(path, a.label)
        with open(tmp_json, "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False)
        shutil.copy2(path, tmp_copy)
    except BaseException:
        for p in (tmp_json, tmp_copy):
            if os.path.exists(p):
                os.remove(p)
        raise
    if os.path.exists(base):
        h = os.path.join(d, "history")
        os.makedirs(h, exist_ok=True)
        shutil.copy2(base, os.path.join(h, "base_%s.json" % stamp()))
    os.replace(tmp_json, base)
    os.replace(tmp_copy, os.path.join(d, copy))
    print("기준본 저장: %s" % base)
    return 0


def cmd_status(a):
    path = os.path.abspath(a.file)
    base = os.path.join(store_dir(path), "base.json")
    if not os.path.exists(base):
        print("기준본 없음 - 먼저 snap")
        return 2
    with open(base, encoding="utf-8") as f:
        meta = json.load(f)["meta"]
    same = meta.get("sha1") == sha1(path)
    print("기준본 %s (%s) · 현재 파일 %s" % (meta.get("made"), meta.get("label") or "", "같음" if same else "바뀜 - diff 필요"))
    return 0 if same else 1


def cmd_diff(a):
    path = os.path.abspath(a.file)
    d = store_dir(path)
    base = a.base or os.path.join(d, "base.json")
    if not os.path.exists(base):
        print("기준본 없음: %s - 먼저 snap" % base)
        return 2
    res = compare(read_snap(base), snapshot(path), _ignore(a.ignore))
    attach(res, a.code or [], a.code[0] if a.code else os.path.dirname(path))
    md = markdown(res)
    t = stamp()
    with open(os.path.join(d, "diff_%s.json" % t), "w", encoding="utf-8") as f:
        json.dump(res, f, ensure_ascii=False, indent=1)
    with open(os.path.join(d, "diff_%s.md" % t), "w", encoding="utf-8") as f:
        f.write(md)
    print(md)
    print("저장: %s (.md·.json) · 총 %d건" % (os.path.join(d, "diff_%s" % t), res["summary"]["total"]))
    return 0


def cmd_verify(a):
    res = compare(read_snap(os.path.abspath(a.expected)), read_snap(os.path.abspath(a.actual)), _ignore(a.ignore))
    if res["summary"]["total"] == 0:
        print("일치: 코드로 다시 만든 문서가 기대 문서와 같음(무시: %s)" % (a.ignore or "없음"))
        return 0
    print(markdown(res, title="남은 차이(기대 -> 실제)"))
    return 1


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description="Word·PPT 수정 추적·코드 반영 도구")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("snap")
    p.add_argument("file")
    p.add_argument("--label", default="")
    p = sub.add_parser("status")
    p.add_argument("file")
    p = sub.add_parser("diff")
    p.add_argument("file")
    p.add_argument("--base")
    p.add_argument("--code", nargs="*")
    p.add_argument("--ignore", default="")
    p = sub.add_parser("verify")
    p.add_argument("expected")
    p.add_argument("actual")
    p.add_argument("--ignore", default="")
    a = ap.parse_args()
    return {"snap": cmd_snap, "status": cmd_status, "diff": cmd_diff, "verify": cmd_verify}[a.cmd](a) or 0


if __name__ == "__main__":
    sys.exit(main())
