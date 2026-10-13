# -*- coding: utf-8 -*-
"""xlsync_excel - 엑셀(COM)로 사본을 한 번 열어 저장해 '엑셀이 저장한 모양'으로 맞춘다(원본은 건드리지 않음).

openpyxl 로 만든 파일과 엑셀에서 저장한 파일은 내용이 같아도 기본 글꼴·열너비 반올림 같은 저장 흔적이 다르다.
기준본(snap)과 재생성본(verify)을 엑셀로 한 번 저장한 사본으로 읽으면 사용자가 엑셀에서 저장한 파일과 같은 조건이 된다.
사용자가 열어 둔 엑셀과 섞이지 않게 별도 인스턴스(DispatchEx)를 쓰고, 매크로·이벤트는 끈다.
"""
import os
import re
import shutil
import tempfile
import zipfile


def excel_saved(path):
    """엑셀이 마지막으로 저장한 파일인지(workbook.xml 의 엑셀 전용 표식)."""
    try:
        with zipfile.ZipFile(path) as z:
            wbx = z.read("xl/workbook.xml").decode("utf-8", "replace")
        return bool(re.search(r"xr:revisionPtr|x15ac:absPath|<fileVersion[^>]*appName=\"xl\"", wbx))
    except Exception:
        return False


def available():
    try:
        import win32com.client  # noqa: F401
        return True
    except Exception:
        return False


def resave_copy(path, workdir=None):
    """path 를 임시 폴더에 복사해 엑셀로 열고 저장한 사본 경로. 엑셀이 없으면 None."""
    if not available():
        return None
    import pythoncom
    import win32com.client
    d = workdir or tempfile.mkdtemp(prefix="xlsync_")
    dst = os.path.join(d, "resaved_" + os.path.basename(path))
    shutil.copy2(path, dst)
    pythoncom.CoInitialize()
    xl = None
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.EnableEvents = False
        try:
            xl.AutomationSecurity = 3          # msoAutomationSecurityForceDisable - 열 때 매크로 실행 안 함
        except Exception:
            pass
        wb = xl.Workbooks.Open(os.path.abspath(dst), 0, False)
        wb.Save()
        wb.Close(False)
    finally:
        if xl is not None:
            xl.Quit()
        pythoncom.CoUninitialize()
    return dst


def normalized(path, workdir=None, force=False):
    """엑셀 저장본이면 그대로, 아니면(또는 force) 엑셀로 저장한 사본. 엑셀이 없으면 원본."""
    if not force and excel_saved(path):
        return path, False
    cp = resave_copy(path, workdir)
    return (cp, True) if cp else (path, False)
