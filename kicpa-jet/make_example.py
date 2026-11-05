# -*- coding: utf-8 -*-
"""가상 회사 예시 자료 - 고객 자료 없이 팩 전체를 돌려 보는 용도(배포 · 강의).

  python -X utf8 make_example.py [--out <폴더>]   (기본: 이 스킬 폴더의 examples)

만드는 것(회사마다): 더존 표준 머리글 분개장(xlsx) · TB(csv, 계정코드 · 계정과목 · 구분 · 기초 · 기말)
  가나상사(가상) - K-IFRS 도소매, 2026 회계연도, 12월 31일 마감 대체 전표 포함
  다라정밀(가상) - 일반기업회계기준 제조, 월말 원가 대체(재공품 · 제품) 전표 포함
분개장과 TB 는 서로 맞게 만든다(완전성 차이 0). 시나리오가 걸리도록 몇 전표를 일부러 심는다(주말 · 기말 고액 · 끝자리 0 ·
키워드 · 적요 없음 · 매출 · 대여금 비정상 상대계정 · 드문 계정 · 거래처 없는 채권). 난수 시드 고정 - 매번 같은 자료.
"""
import argparse
import csv
import datetime as dt
import os
import random
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from kicpa_core import THEME_DIR  # noqa: E402,F401  (excel-theme 경로)
from openpyxl import Workbook  # noqa: E402

NAMES = {101: "현금", 103: "보통예금", 108: "외상매출금", 109: "대손충당금", 114: "단기대여금", 120: "미수금", 131: "선급금",
         135: "부가세대급금", 146: "상품", 150: "제품", 153: "원재료", 169: "재공품", 179: "장기대여금", 202: "건물",
         203: "감가상각누계액", 206: "기계장치", 207: "감가상각누계액", 212: "비품", 213: "감가상각누계액", 251: "외상매입금",
         253: "미지급금", 254: "예수금", 255: "부가세예수금", 259: "선수금", 260: "단기차입금", 293: "장기차입금",
         295: "퇴직급여충당부채", 331: "자본금", 375: "이월이익잉여금", 377: "미처분이익잉여금", 400: "손익", 401: "상품매출",
         404: "제품매출", 451: "상품매출원가", 455: "제품매출원가", 501: "원재료비", 504: "임금", 518: "감가상각비",
         801: "급여", 811: "복리후생비", 818: "감가상각비", 819: "지급임차료", 831: "지급수수료", 901: "이자수익",
         930: "잡이익", 931: "이자비용", 961: "재해손실", 998: "법인세등"}
KIND = lambda c: "자산" if c < 251 else ("부채" if c < 331 else ("자본" if c < 400 else ("제조" if 500 <= c < 800 else "손익")))
CUST = ["가람상회", "누리마트", "다온유통", "라온식품", "마루상사", "바다물산"]
VEND = ["한빛산업", "새솔물류", "온누리공업", "해솔기업"]


class Book:
    def __init__(self):
        self.v = []          # (일자, 번호, [(코드, 차, 대, 거래처, 적요)])
        self.no = {}

    def add(self, d, lines, memo=None):
        n = self.no.get(d, 0) + 1
        self.no[d] = n
        self.v.append((d, n, [(c, dr, cr, cu, memo if m is None else m) for c, dr, cr, cu, m in lines]))

    def lines(self):
        for d, n, L in sorted(self.v, key=lambda x: (x[0], x[1])):
            for c, dr, cr, cu, m in L:
                yield d, n, c, dr, cr, cu, m


def L(c, dr=0, cr=0, cu="", m=None):
    return (c, dr, cr, cu, m)


def month_end(y, m):
    return dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)


def biz_day(y, m, d):
    x = dt.date(y, m, d)
    while x.weekday() >= 5:
        x += dt.timedelta(days=1)
    return x


def trade(seed, mfg):
    rnd = random.Random(seed)
    y = 2026
    B = Book()
    rev, cogs = (404, 455) if mfg else (401, 451)
    for m in range(1, 13):
        for k in range(18 if not mfg else 17):              # 매출
            d = biz_day(y, m, rnd.randint(1, 26))
            amt = rnd.randrange(20, 90) * 1000000 + rnd.randrange(1, 999) * 1000 + rnd.randrange(1, 9) * 100
            vat = round(amt * 0.1)
            cu = rnd.choice(CUST)
            B.add(d, [L(108, amt + vat, 0, cu), L(rev, 0, amt, cu), L(255, 0, vat, cu)], "%s 매출" % cu)
        for k in range(12):                                  # 회수
            d = biz_day(y, m, rnd.randint(5, 27))
            amt = rnd.randrange(50, 110) * 1000000 + rnd.randrange(1, 999) * 1000
            cu = rnd.choice(CUST)
            B.add(d, [L(103, amt, 0, cu), L(108, 0, amt, cu)], "외상대 회수")
        buy = 153 if mfg else 146
        for k in range(10):                                  # 매입
            d = biz_day(y, m, rnd.randint(1, 25))
            amt = rnd.randrange(30, 95) * 1000000 + rnd.randrange(1, 999) * 1000 + 300
            vat = round(amt * 0.1)
            ve = rnd.choice(VEND)
            B.add(d, [L(buy, amt, 0, ve), L(135, vat, 0, ve), L(251, 0, amt + vat, ve)], "%s 매입" % ve)
        for k in range(8):                                   # 지급
            d = biz_day(y, m, rnd.randint(5, 28))
            amt = rnd.randrange(60, 120) * 1000000 + rnd.randrange(1, 999) * 1000
            ve = rnd.choice(VEND)
            B.add(d, [L(251, amt, 0, ve), L(103, 0, amt, ve)], "외상대 지급")
        pay = biz_day(y, m, 25)                              # 급여
        if mfg:
            B.add(pay, [L(504, 180000000), L(801, 90000000), L(254, 0, 27000000), L(103, 0, 243000000)], "%d월 급여" % m)
        else:
            B.add(pay, [L(801, 150000000), L(254, 0, 15000000), L(103, 0, 135000000)], "%d월 급여" % m)
        B.add(biz_day(y, m, 10), [L(254, 27000000 if mfg else 15000000), L(103, 0, 27000000 if mfg else 15000000)], "예수금 납부")
        B.add(biz_day(y, m, 20), [L(811, 8350000), L(103, 0, 8350000)], "식대 · 경조사비")
        B.add(biz_day(y, m, 5), [L(819, 12000000), L(135, 1200000), L(103, 0, 13200000, "한결빌딩")], "사무실 임차료")
        B.add(biz_day(y, m, 15), [L(831, 3300000), L(103, 0, 3300000)], "세무 · 전산 수수료")
        e = month_end(y, m)
        if mfg:
            B.add(e, [L(518, 25000000), L(207, 0, 25000000)], "기계장치 감가상각")
            B.add(e, [L(818, 4000000), L(213, 0, 4000000)], "비품 감가상각")
            mat = rnd.randrange(400, 520) * 1000000
            B.add(e, [L(169, mat + 180000000 + 25000000), L(501, 0, mat), L(504, 0, 180000000), L(518, 0, 25000000)], "원가 대체")
            B.add(e, [L(501, mat), L(153, 0, mat)], "원재료 출고")
            fg = mat + 205000000
            B.add(e, [L(150, fg), L(169, 0, fg)], "제품 입고")
            cs = round(fg * 0.97)
            B.add(e, [L(cogs, cs), L(150, 0, cs)], "제품매출원가 대체")
        else:
            B.add(e, [L(818, 6000000), L(203, 0, 2000000), L(213, 0, 4000000)], "감가상각")
            cs = rnd.randrange(690, 760) * 1000000
            B.add(e, [L(cogs, cs), L(146, 0, cs)], "상품매출원가 대체")
        if m % 3 == 0:
            B.add(e, [L(931, 9750000), L(103, 0, 9750000, "가온은행")], "차입금 이자")
            B.add(e, [L(103, 1240000, 0, "가온은행"), L(901, 0, 1240000, "가온은행")], "예금 이자")
    # ── 심은 시나리오 ──
    B.add(dt.date(y, 3, 14), [L(253, 0, 66000000, "새솔물류"), L(831, 60000000), L(135, 6000000)], "물류 컨설팅")   # 토요일
    B.add(dt.date(y, 12, 31), [L(108, 550000000, 0, "누리마트"), L(rev, 0, 500000000, "누리마트"), L(255, 0, 50000000, "누리마트")], "연말 일괄 매출")
    B.add(dt.date(y, 6, 30), [L(114, 300000000, 0, "대표이사"), L(103, 0, 300000000)], "대표이사 대여")
    B.add(dt.date(y, 8, 12), [L(103, 1200000), L(930, 0, 1200000)], None)                            # 적요 없음
    B.add(dt.date(y, 9, 21), [L(179, 200000000, 0, "다온유통"), L(rev, 0, 200000000, "다온유통")], "매출 대체 정정")  # 매출 비정상 상대계정 · 키워드
    B.add(dt.date(y, 10, 15), [L(114, 80000000, 0, "마루상사"), L(253, 0, 80000000, "마루상사")], "단기 자금 대여")  # 대여금 상대계정
    B.add(dt.date(y, 11, 3), [L(961, 15000000), L(212, 0, 15000000)], "창고 침수 비품 폐기")           # 드문 계정
    B.add(dt.date(y, 7, 9), [L(108, 33000000, 0, ""), L(rev, 0, 30000000, ""), L(255, 0, 3000000, "")], "현장 판매")  # 거래처 없음
    B.add(dt.date(y, 12, 31), [L(998, 120000000), L(253, 0, 120000000, "관할세무서")], "법인세 추산")
    return B


OPEN = {
    "가나상사": {101: 8000000, 103: 1850000000, 108: 2400000000, 109: -24000000, 146: 1650000000, 202: 3200000000,
              203: -640000000, 212: 240000000, 213: -96000000, 251: -1450000000, 253: -180000000, 254: -15000000,
              260: -1500000000, 331: -1000000000, 375: 0, 377: -4443000000},
    "다라정밀": {101: 5000000, 103: 2200000000, 108: 3100000000, 109: -31000000, 153: 900000000, 150: 1200000000,
              169: 300000000, 202: 2800000000, 203: -700000000, 206: 4800000000, 207: -1800000000, 212: 180000000,
              213: -60000000, 251: -1900000000, 253: -240000000, 254: -27000000, 293: -2500000000, 295: -620000000,
              331: -2000000000, 377: -5607000000},
}


def close_year(B, y):
    """12.31 마감 대체 전표(더존): 수익 · 비용 → 400 손익 → 377 미처분이익잉여금."""
    bal = {}
    for d, n, c, dr, cr, cu, m in B.lines():
        if KIND(c) == "손익":
            bal[c] = bal.get(c, 0) + dr - cr
    e = dt.date(y, 12, 31)
    rev = [L(c, -v, 0) for c, v in bal.items() if v < 0]
    exp = [L(c, 0, v) for c, v in bal.items() if v > 0]
    B.add(e, rev + [L(400, 0, -sum(v for v in bal.values() if v < 0))], "손익 대체")
    B.add(e, [L(400, sum(v for v in bal.values() if v > 0))] + exp, "손익 대체")
    ni = -sum(bal.values())
    B.add(e, [L(400, ni), L(377, 0, ni)], "당기순이익 대체")
    return bal


def write(company, B, opening, outdir, closing=True):
    os.makedirs(outdir, exist_ok=True)
    pl = close_year(B, 2026) if closing else None
    wb = Workbook()
    ws = wb.active
    ws.title = "분개장"
    ws.append(["분개장"])
    ws.append(["회사명 : %s(가상)" % company, None, None, "2026년 1월 1일 ~ 2026년 12월 31일"])
    ws.append([])
    ws.append(["일자", "번호", "계정코드", "계정과목", "차변", "대변", "거래처", "적요"])
    tdr = tcr = 0
    for d, n, c, dr, cr, cu, m in B.lines():
        ws.append([d, n, c, NAMES[c], dr or None, cr or None, cu or None, m or None])
        tdr += dr
        tcr += cr
    ws.append(["합계", None, None, None, tdr, tcr])
    wb.properties.creator = "이재현"
    jp = os.path.join(outdir, "%s_분개장_2026.xlsx" % company)
    wb.save(jp)
    mv = {}
    for d, n, c, dr, cr, cu, m in B.lines():
        mv[c] = mv.get(c, 0) + dr - cr
    # TB: 재무상태표 = 마감 후, 손익 = 마감 전 금액(더존 손익계산서), 제조 = 원가 대체 전(제조원가명세서), 400 손익은 0 이라 뺌
    rows = []
    rnd = random.Random(len(company))
    codes = sorted(set(opening) | set(mv))
    for c in codes:
        if c == 400:
            continue
        k = KIND(c)
        if k in ("자산", "부채", "자본"):
            rows.append([c, NAMES[c], k, opening.get(c, 0), opening.get(c, 0) + mv.get(c, 0)])
        elif k == "손익":                            # 기초 칸 = 전기 금액(가상 - 당기의 90~105%)
            amt = pl.get(c, 0) if pl else mv.get(c, 0)
            rows.append([c, NAMES[c], k, round(amt * rnd.uniform(0.9, 1.05), -3), amt])
        else:                                       # 제조 = 원가 대체 전 금액(제조원가명세서)
            amt = sum(dr - cr for d, n, cc, dr, cr, cu, m in B.lines() if cc == c and m != "원가 대체")
            rows.append([c, NAMES[c], k, round(amt * rnd.uniform(0.9, 1.05), -3), amt])
    tp = os.path.join(outdir, "%s_TB_2026.csv" % company)
    with open(tp, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["계정코드", "계정과목", "구분", "기초", "기말"])   # 손익 · 제조 계정의 기초 = 전기 금액
        w.writerows(rows)
    return jp, tp, len(list(B.lines()))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(HERE, "examples"))
    a = ap.parse_args()
    for company, mfg, seed in (("가나상사", False, 11), ("다라정밀", True, 23)):
        B = trade(seed, mfg)
        jp, tp, n = write(company, B, OPEN[company], a.out)
        print("%s: 분개장 %d줄 → %s · TB → %s" % (company, n, jp, tp))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
