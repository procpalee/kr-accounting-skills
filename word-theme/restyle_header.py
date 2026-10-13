# -*- coding: utf-8 -*-
"""기존 .docx 의 제목부를 현행 block 스타일(정보표 2×4: 회사명│결산일 / 조서번호│작성자 + 제목 + 하단선)로 교체. 본문은 건드리지 않는다.
원본은 <이름>.bak.docx 로 백업(이미 있으면 유지).

인식하는 기존 제목부:
  ① meta 스타일  — 첫 표(1행 2열, 우측 셀에 정보표) + 하단선 문단
  ② 구 block     — 첫 표(라벨·값 격자) + 제목 문단 + 하단선 문단
  ③ 현행 block   — 제목 문단 + 첫 표(2행 6열 밑줄형) + 빈 문단   (재적용 시)

사용: python restyle_header.py <file.docx> [--no 6002] [--company ...] [--fye 2026.12.31] [--author ...]
      (지정한 값이 기존 값보다 우선. 검토자·작성일자 등 구 항목은 버린다 — 2026-09-07 확정 4항목)
"""
import sys, os, shutil, argparse, re
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.oxml.ns import qn
from word_theme import _title_block, get_theme

ALIASES = {"회사": "회사명", "기준일": "결산일", "관련계정": None, "관련 계정": None, "관련 조서": None,
           "작성일": "작성일자", "검토일": "검토일자", "작성": "작성자", "검토": "검토자"}


def _txt(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()


def _cells_text(tbl_el):
    """표 요소 → 행별 셀 텍스트(중첩 표 제외)."""
    rows = []
    for tr in tbl_el.findall(qn("w:tr")):
        cells = []
        for tc in tr.findall(qn("w:tc")):
            texts = []
            for p in tc.findall(qn("w:p")):
                texts.append("".join(t.text or "" for t in p.iter(qn("w:t"))))
            cells.append("\n".join(texts).strip())
        rows.append(cells)
    return rows


def split_name_date(v):
    """'이재현 / 2026.09.04' → ('이재현', '2026.09.04'); '이재현' → ('이재현', '')."""
    parts = [x.strip() for x in re.split(r"\s*/\s*", (v or "").strip())]
    name = parts[0] if parts else ""
    date = parts[1] if len(parts) > 1 else ""
    if re.fullmatch(r"\d{4}[.\-년]\s*\d{1,2}[.\-월]\s*\d{1,2}[.\-일]?\.?", name):
        return "", name
    return name, date


def parse_pairs(rows):
    meta = {}
    for ri, r in enumerate(rows):
        for i in range(0, len(r) - 1, 2):
            k, v = r[i].strip(), r[i + 1].strip()
            if not k:
                continue
            if k == "일자":
                k = "작성일자" if ri == 0 else "검토일자"
            k = ALIASES.get(k, k)
            if k is None:
                continue
            meta[k] = v
    return meta


def detect(doc):
    """(kind, elements_to_remove, title, meta)"""
    body = doc.element.body
    kids = [e for e in body if e.tag != qn("w:sectPr")]
    if not kids:
        return None
    first = kids[0]
    # ③-a v2 block: 제목 문단 + 표(2×6) + 빈 문단
    if first.tag == qn("w:p") and len(kids) > 1 and kids[1].tag == qn("w:tbl"):
        rows = _cells_text(kids[1])
        if len(rows) == 2 and len(rows[0]) == 6:
            rm = [kids[0], kids[1]]
            if len(kids) > 2 and kids[2].tag == qn("w:p") and not _txt(kids[2]):
                rm.append(kids[2])
            return "block-v2", rm, _txt(kids[0]), parse_pairs(rows)
    # ③-b v3 block: 표(2×6) + 제목 문단
    if first.tag == qn("w:tbl"):
        rows = _cells_text(first)
        if len(rows) == 2 and len(rows[0]) == 6 and len(kids) > 1 and kids[1].tag == qn("w:p"):
            return "block-v3", [kids[0], kids[1]], _txt(kids[1]), parse_pairs(rows)
    if first.tag != qn("w:tbl"):
        return None
    rows = _cells_text(first)
    trs = first.findall(qn("w:tr"))
    # ① meta: 1행 2열, 우측 셀에 내부 표
    if len(trs) == 1 and len(trs[0].findall(qn("w:tc"))) == 2:
        left, right = trs[0].findall(qn("w:tc"))
        title = "".join(t.text or "" for t in left.findall(qn("w:p"))[0].iter(qn("w:t"))).strip()
        inner = right.find(qn("w:tbl"))
        meta = parse_pairs(_cells_text(inner)) if inner is not None else {}
        rm = [first]
        nxt = first.getnext()
        if nxt is not None and nxt.tag == qn("w:p") and not _txt(nxt):
            rm.append(nxt)
        return "meta", rm, title, meta
    # ② 구 block: 라벨·값 격자(4열) + 제목 문단 + 하단선 문단
    if rows and len(rows[0]) == 4:
        meta = parse_pairs(rows)
        rm = [first]
        nxt = first.getnext()
        title = ""
        # 표와 제목 사이의 빈 문단은 함께 제거하고 제목은 첫 비어있지 않은 문단에서 취함
        while nxt is not None and nxt.tag == qn("w:p") and not _txt(nxt):
            rm.append(nxt); nxt = nxt.getnext()
        if nxt is not None and nxt.tag == qn("w:p"):
            title = _txt(nxt); rm.append(nxt)
            nxt2 = nxt.getnext()
            if nxt2 is not None and nxt2.tag == qn("w:p") and not _txt(nxt2):
                rm.append(nxt2)
        return "oldblock", rm, title, meta
    return None


def restyle(path, no=None, company=None, author=None, fye=None, **_ignored):
    bak = os.path.splitext(path)[0] + ".bak.docx"
    if not os.path.exists(bak):
        shutil.copy2(path, bak)
    doc = Document(path)
    found = detect(doc)
    if not found:
        return "인식 가능한 제목부가 없음 — 변경 없음"
    kind, rm, title, old = found
    a_name, a_date = split_name_date(old.get("작성자", ""))
    r_name, r_date = split_name_date(old.get("검토자", ""))
    meta = {
        "회사명": company or old.get("회사명", ""),
        "결산일": fye or old.get("결산일", ""),
        "조서번호": no if no is not None else old.get("조서번호", ""),
        "작성자": author or a_name,
    }
    body = doc.element.body
    for e in rm:
        body.remove(e)
    th = get_theme("default")
    before = list(body)
    _title_block(doc, th, title, None, None, meta)
    new_elems = [e for e in list(body) if e not in before and e.tag != qn("w:sectPr")]
    for e in new_elems:
        body.remove(e)
    for e in reversed(new_elems):
        body.insert(0, e)
    doc.save(path)
    return f"[{kind}] 제목부 교체 — {meta}"


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    for k in ("no", "company", "author", "fye"):
        ap.add_argument("--" + k, dest=k)
    a = ap.parse_args()
    print(os.path.basename(a.file) + ": " + restyle(a.file, a.no, a.company, a.author, a.fye))
