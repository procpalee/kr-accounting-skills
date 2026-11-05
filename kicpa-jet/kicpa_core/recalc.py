# -*- coding: utf-8 -*-
"""LibreOffice(헤드리스) 재계산 - 조서 사본을 다시 계산해 수식 결과값을 읽는다(원본은 읽기만)."""
import glob
import os
import shutil
import subprocess
import tempfile

SOFFICE = [r"C:\Program Files\LibreOffice\program\soffice.exe", r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
           "/usr/bin/soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
ERR = ("#REF!", "#NAME?", "#VALUE!", "#N/A", "#DIV/0!", "#NUM!", "#NULL!", "Err:")


def soffice():
    for p in SOFFICE:
        if os.path.exists(p):
            return p
    p = shutil.which("soffice")
    if p:
        return p
    raise SystemExit("LibreOffice(soffice)를 찾지 못했습니다 - 설치 후 다시 실행")


def recalc(path, timeout=3600):
    tmp = tempfile.mkdtemp(prefix="kicpa_")
    src = os.path.join(tmp, "in.xlsx")
    shutil.copy(path, src)
    out = os.path.join(tmp, "out")
    os.makedirs(out)
    prof = "file:///" + os.path.join(tmp, "profile").replace("\\", "/")
    subprocess.run([soffice(), "-env:UserInstallation=" + prof, "--headless", "--calc", "--convert-to", "xlsx",
                    "--outdir", out, src], capture_output=True, timeout=timeout)
    res = glob.glob(os.path.join(out, "*.xlsx"))
    if not res:
        shutil.rmtree(tmp, ignore_errors=True)
        raise SystemExit("LibreOffice 재계산 실패")
    return res[0], tmp


def error_cells(wb, skip=()):
    bad = []
    for ws in wb.worksheets:
        if ws.title in skip:
            continue
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith(ERR):
                    bad.append("%s!%s" % (ws.title, c.coordinate))
    return bad
