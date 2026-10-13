# -*- coding: utf-8 -*-
"""기존 .docx 의 표 테두리 굵기를 일반 굵기(0.5pt)로 통일 — 본문·문구는 건드리지 않는다.

사용: python restyle_tables.py <file.docx> [<file2.docx> ...]
      원본은 같은 폴더에 <이름>.bak.docx 로 백업한다.
"""
import sys, shutil, os
from docx import Document
from docx.oxml.ns import qn

TARGET_SZ = "4"   # 1/8pt 단위 → 0.5pt


def restyle(path):
    bak = os.path.splitext(path)[0] + ".bak.docx"
    if not os.path.exists(bak):
        shutil.copy2(path, bak)
    doc = Document(path)
    n = 0
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                tcPr = cell._tc.tcPr
                if tcPr is None:
                    continue
                borders = tcPr.find(qn("w:tcBorders"))
                if borders is None:
                    continue
                for b in borders:
                    if b.get(qn("w:val")) not in (None, "nil", "none") and b.get(qn("w:sz")) not in (None, TARGET_SZ):
                        b.set(qn("w:sz"), TARGET_SZ)
                        n += 1
    doc.save(path)
    return n


if __name__ == "__main__":
    for f in sys.argv[1:]:
        print(f"{os.path.basename(f)}: {restyle(f)} borders → 0.5pt (backup .bak.docx)")
