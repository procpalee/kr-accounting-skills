# -*- coding: utf-8 -*-
"""분개장 읽기 - 머리글 이름으로 열을 찾는다(공백 · 줄바꿈 무시).

두 형식을 읽는다.
  더존 표준   일자 · 번호 · 계정코드 · 계정과목 · 차변 · 대변 · 적요 · 거래처 … (인쇄 미리보기 [엑셀], 사용자정의 출력)
  ERP 내보내기 승인일 · 승인번호 · 차변금액 · 차변계정과목 · 대변계정과목 · 대변금액 · 적요 · 거래처 · 작성일
              (계정은 '[1020000]당좌예금' 처럼 코드와 이름이 한 칸, 한 줄에 차변 또는 대변 한쪽만)
한 기간이 여러 파일로 나뉘어 있으면(반월 · 월 단위) parse_many 로 합친다. 원본 파일은 읽기만 한다.
"""
import datetime as dt
import glob
import os
import re

from . import KicpaError

HEAD = {
    "일자": ("일자", "전표일자", "회계일자", "년/월/일", "승인일", "승인일자", "기표일", "기표일자", "전기일"),
    "전표번호": ("전표번호", "번호", "승인번호", "전표No", "전표NO"),
    "계정코드": ("계정코드", "계정과목코드"),
    "계정과목": ("계정과목", "계정과목명", "계정명"),
    "차변계정": ("차변계정과목", "차변계정"),
    "대변계정": ("대변계정과목", "대변계정"),
    "차변": ("차변", "차변금액"),
    "대변": ("대변", "대변금액"),
    "적요": ("적요", "적요명"),
    "거래처코드": ("거래처코드",),
    "거래처명": ("거래처명", "거래처"),
    "구분": ("구분",),
    "유형": ("유형", "전표유형"),
    "부서사원": ("부서/사원명", "부서사원명", "사원명"),
    "작성일": ("작성일", "작성일자", "입력일", "입력일자"),
    "입력자": ("입력자", "작성자"),
    "승인자": ("승인자",),
}
PERIOD = re.compile(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*~\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일")
ACCT = re.compile(r"^\s*\[\s*([0-9A-Za-z]+)\s*\]\s*(.*)$")


def norm(v):
    return re.sub(r"\s+", "", str(v if v is not None else ""))


def to_num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        x = float(s)
    except ValueError:
        return 0.0
    return -x if neg else x


def code_of(v):
    """계정코드 칸 → int(숫자 코드) 또는 문자열."""
    s = str(v if v is not None else "").strip()
    if re.fullmatch(r"\d+(\.0)?", s):
        return int(float(s))
    return s


def split_acct(v):
    """'[1020000]당좌예금(OO은행)' → (1020000, '당좌예금(OO은행)'). 대괄호가 없으면 (None, 이름)."""
    s = str(v if v is not None else "").strip()
    m = ACCT.match(s)
    if m:
        return code_of(m.group(1)), m.group(2).strip()
    return None, s


def load_rows(path):
    """첫 시트 전체 → 2차원 리스트(xls 날짜 칸은 datetime)."""
    low = path.lower()
    if low.endswith((".xlsx", ".xlsm")):
        import openpyxl
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        ws.reset_dimensions()                      # 더존 xlsx 는 dimension 이 A1 로 적혀 있어 read_only 가 한 줄만 읽음
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    if low.endswith(".csv"):
        import csv
        for enc in ("utf-8-sig", "cp949"):
            try:
                with open(path, encoding=enc, newline="") as f:
                    return [list(r) for r in csv.reader(f)]
            except UnicodeDecodeError:
                continue
    import xlrd
    try:
        wb = xlrd.open_workbook(path)
    except xlrd.biffh.XLRDError:
        return _html_rows(path)                     # 확장자만 xls 인 HTML 표(일부 ERP 내보내기)
    ws = wb.sheet_by_index(0)
    out = []
    for r in range(ws.nrows):
        row = []
        for c in range(ws.ncols):
            cell = ws.cell(r, c)
            row.append(xlrd.xldate_as_datetime(cell.value, wb.datemode) if cell.ctype == xlrd.XL_CELL_DATE else cell.value)
        out.append(row)
    return out


def _html_rows(path):
    from html.parser import HTMLParser
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8", "cp949"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None or "<t" not in text.lower():
        raise KicpaError("엑셀 · HTML 어느 쪽으로도 읽지 못했습니다: %s" % os.path.basename(path))

    class P(HTMLParser):
        def __init__(self):
            super().__init__()
            self.rows, self.row, self.cell, self.on = [], None, [], False

        def handle_starttag(self, tag, a):
            if tag == "tr":
                self.row = []
            elif tag in ("td", "th"):
                self.on, self.cell = True, []

        def handle_endtag(self, tag):
            if tag in ("td", "th") and self.row is not None:
                self.row.append("".join(self.cell).strip())
                self.on = False
            elif tag == "tr" and self.row is not None:
                self.rows.append(self.row)
                self.row = None

        def handle_data(self, d):
            if self.on:
                self.cell.append(d)

    p = P()
    p.feed(text)
    return p.rows


def _date(v, start=None):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})$", s)
    if m and start:
        mm, dd = int(m.group(1)), int(m.group(2))
        y = start.year if (mm, dd) >= (start.month, start.day) else start.year + 1
        return dt.date(y, mm, dd)
    m = re.match(r"^(\d{4})[-./]?(\d{1,2})[-./]?(\d{1,2})", s)
    if m:
        try:
            return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def _find_header(rows):
    for i, r in enumerate(rows[:60]):
        cells = {norm(x) for x in r if x not in (None, "")}
        if cells & set(HEAD["일자"]) and cells & (set(HEAD["차변"]) | set(HEAD["차변계정"])):
            return i
    return None


def parse_journal(path, year=None):
    """분개장 파일 하나 → (lines, info). lines 는 파일 안 순서, info = {파일, 시작, 끝, 원본합계, 열, 형식}."""
    rows = load_rows(path)
    hdr = _find_header(rows)
    if hdr is None:
        raise KicpaError("분개장 머리글(일자 · 차변 · 대변)을 찾지 못했습니다: %s" % os.path.basename(path))
    col = {}
    for j, h in enumerate(rows[hdr]):
        n = norm(h)
        for f, names in HEAD.items():
            if f not in col and n in names:
                col[f] = j
    erp = "차변계정" in col and "대변계정" in col
    need = ("일자", "전표번호", "차변", "대변") + (() if erp else ("계정과목",))
    miss = [f for f in need if f not in col]
    if miss:
        raise KicpaError("분개장에 필요한 열이 없습니다(%s): %s" % (os.path.basename(path), miss))
    text = " ".join(str(x) for r in rows[:hdr] for x in r if x)
    m = PERIOD.search(text)
    start = end = None
    if m:
        start = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        end = dt.date(int(m.group(4)), int(m.group(5)), int(m.group(6)))
    elif year:
        start = dt.date(int(year), 1, 1)
    g = lambda r, f: r[col[f]] if f in col and col[f] < len(r) else None
    lines, total = [], None
    fname = os.path.basename(path)
    for i, r in enumerate(rows[hdr + 1:]):
        if not r or all(x in (None, "") for x in r):
            continue
        first = norm(g(r, "일자"))
        if first in ("합계", "총계") or norm(r[0]) in ("합계", "총계"):
            total = (to_num(g(r, "차변")), to_num(g(r, "대변")))
            continue
        d = _date(g(r, "일자"), start)
        if d is None:
            continue
        vno = g(r, "전표번호")
        vno = str(int(vno)) if isinstance(vno, float) else str(vno or "").strip()
        dr, cr = to_num(g(r, "차변")), to_num(g(r, "대변"))
        if erp:
            side = g(r, "차변계정") if (dr or not g(r, "대변계정")) else g(r, "대변계정")
            code, name = split_acct(side)
        else:
            code = code_of(g(r, "계정코드")) if "계정코드" in col else None
            c2, name = split_acct(g(r, "계정과목"))
            if code in (None, ""):
                code = c2
        if code in (None, ""):
            code = name
        wd = _date(g(r, "작성일")) if "작성일" in col else None
        lines.append({"i": i, "파일": fname, "일자": d, "전표번호": vno.zfill(5),
                      "키": "%s-%s" % (d.isoformat(), vno.zfill(5)), "코드": code, "계정과목": norm(name),
                      "차변": dr, "대변": cr, "적요": str(g(r, "적요") or "").strip(),
                      "거래처코드": str(g(r, "거래처코드") or "").strip(), "거래처명": str(g(r, "거래처명") or "").strip(),
                      "구분": str(g(r, "구분") or "").strip(), "유형": str(g(r, "유형") or "").strip(),
                      "부서사원": str(g(r, "부서사원") or "").strip(), "작성일": wd,
                      "입력자": str(g(r, "입력자") or "").strip(), "승인자": str(g(r, "승인자") or "").strip()})
    if not lines:
        raise KicpaError("분개장에 전표 줄이 없습니다: %s" % fname)
    if start is None:
        start = dt.date(min(x["일자"] for x in lines).year, 1, 1)
    if end is None:
        end = max(x["일자"] for x in lines)
    fields = {f: (f in col) for f in HEAD}
    if erp:
        fields["계정코드"] = fields["계정과목"] = True
    return lines, {"파일": fname, "시작": start, "끝": end, "원본합계": total, "열": fields,
                   "형식": "ERP 내보내기" if erp else "더존 분개장", "줄": len(lines)}


def expand(paths):
    """파일 · 폴더 · 와일드카드 → 분개장 파일 목록(이름 순, ~$ 임시 파일 · 압축 제외)."""
    out = []
    for p in paths:
        if os.path.isdir(p):
            cand = [f for ext in ("*.xlsx", "*.xls", "*.xlsm", "*.csv") for f in glob.glob(os.path.join(p, ext))]
        else:
            cand = glob.glob(p) or [p]
        out += [f for f in cand if not os.path.basename(f).startswith("~$")]
    seen, res = set(), []
    for f in sorted(out, key=lambda x: _natkey(os.path.basename(x))):
        k = os.path.abspath(f).lower()
        if k not in seen:
            seen.add(k)
            res.append(f)
    if not res:
        raise KicpaError("분개장 파일이 없습니다: %s" % paths)
    return res


def _natkey(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", s)]


def parse_many(paths, start=None, end=None, year=None):
    """여러 분개장 파일 → (lines(전표키 순), info). start · end 를 주면 그 기간 밖 줄은 빼고 건수를 알린다."""
    files = expand(paths)
    lines, infos = [], []
    for f in files:
        L, info = parse_journal(f, year)
        lines += L
        infos.append(info)
    before = len(lines)
    if start:
        lines = [x for x in lines if x["일자"] >= start]
    if end:
        lines = [x for x in lines if x["일자"] <= end]
    for n, x in enumerate(lines):
        x["i"] = n
    lines.sort(key=lambda x: (x["일자"], x["전표번호"], x["i"]))
    fields = {f: any(i["열"][f] for i in infos) for f in HEAD}
    tot = None
    if all(i["원본합계"] for i in infos):
        tot = (sum(i["원본합계"][0] for i in infos), sum(i["원본합계"][1] for i in infos))
    info = {"파일": infos[0]["파일"] if len(infos) == 1 else "%s 외 %d개" % (infos[0]["파일"], len(infos) - 1),
            "파일목록": infos, "시작": start or min(i["시작"] for i in infos), "끝": end or max(x["일자"] for x in lines),
            "원본합계": tot, "열": fields, "형식": infos[0]["형식"], "기간 밖 제외": before - len(lines)}
    return lines, info


def vouchers_of(lines):
    """전표키 순으로 정렬된 lines → [{키, 일자, s, e}] (s·e = lines 인덱스, 끝 포함)."""
    out = []
    for n, x in enumerate(lines):
        if out and out[-1]["키"] == x["키"]:
            out[-1]["e"] = n
        else:
            out.append({"키": x["키"], "일자": x["일자"], "s": n, "e": n})
    return out
