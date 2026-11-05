# -*- coding: utf-8 -*-
"""kicpa-jet 시나리오 규칙 - 조서 Data_전표 수식과 **같은 규칙**을 파이썬으로 한 것.

40 선정분개 테스트의 전표 목록(만들 때 값)을 뽑고, jet_check.py 가 엑셀 재계산 건수와 대조할 때 쓴다.
"""
import datetime as dt
from collections import Counter

from kicpa_core import coa as C

SCN = ["JS01", "JS02", "JS03", "JS04", "JS05", "JS06", "JS07", "JS08", "JS09", "JS10", "JS11"]
PERIOD_END = ("JS02", "JS10")          # 한공회 3600 18(1) 2) 보고기간말 분개 - 나머지는 3) 기간 전체

SCN_INFO = {   # ID: (테스트 구분, 시나리오 명, 근거, 추출 방법)
    "JS01": ("전표 Scanning", "주말·공휴일 기표", "240 A44 - 정상 사업과정 밖",
             "기표일이 토·일요일이거나 공휴일 목록에 있고 전표 차변합계가 최소 전표금액 이상인 전표"),
    "JS02": ("결산전표 검토", "보고기간말 고액 분개", "240 문단 33(a)(ii) - 보고기간말 분개",
             "기준일 전 기간(일) 안에 기표되고 전표 차변합계가 기준금액 이상인 전표"),
    "JS03": ("전표 Scanning", "끝자리 0 고액", "240 A44 - 반올림된 숫자",
             "줄 금액이 최소 금액 이상이고 끝자리 단위로 나누어떨어지는 줄이 있는 전표"),
    "JS04": ("전표 Scanning", "적요 키워드 분개", "240 A44 - 비경상 분개",
             "적요에 키워드 목록의 낱말이 들어 있는 줄이 있는 전표"),
    "JS05": ("전표 Scanning", "적요 없는 분개", "240 A44 - 설명 없는 분개",
             "적요가 빈 줄이 있는 전표"),
    "JS06": ("상대계정 분석", "매출 상대계정", "240 A44 - 비경상 계정 조합",
             "JS06 본계정(방향)이 있는 전표 중 본계정·허용 상대계정 밖의 계정이 있는 전표"),
    "JS07": ("상대계정 분석", "대여금 상대계정", "240 A44 - 비경상 계정 조합",
             "JS07 본계정(방향)이 있는 전표 중 본계정·허용 상대계정 밖의 계정이 있는 전표"),
    "JS08": ("전표 Scanning", "드문 계정 분개", "240 A44 - 거의 사용되지 않는 계정",
             "기간 중 분개장 줄 수가 기준 건수 이하인 계정이 있는 전표"),
    "JS09": ("전표 Scanning", "거래처 없는 채권채무", "240 A44 - 비경상 분개",
             "대상 계정 줄 중 거래처명이 빈 줄이 있는 전표"),
    "JS10": ("결산전표 검토", "기준일 이후 작성", "240 A44 - 결산 과정·이후에 기록된 분개",
             "작성일(입력일)이 기준일보다 늦은 전표 - 분개장에 작성일 열이 있을 때"),
    "JS11": ("전표 Scanning", "작성일·기표일 차이", "240 A44 - 사후 입력 분개",
             "작성일이 기표일보다 기준 일수 이상 늦은 전표 - 분개장에 작성일 열이 있을 때"),
}

KEYWORDS = ["수정", "정정", "취소", "착오", "오류", "일회성", "가지급", "가수금", "대표이사", "임원",
            "특수관계", "리베이트", "보정", "임의"]


def default_params(lines, info, pm, holidays, has_wd):
    codes = {x["코드"] for x in lines}
    names = {}
    for x in lines:
        names.setdefault(x["코드"], x["계정과목"])
    yrs = sorted({info["시작"].year, info["끝"].year})
    hol = sorted(d for y in yrs for d in holidays.get(str(y), []))
    rnd = lambda v, k: max(int(v) // k * k, k)
    return {
        "기준일": info["끝"], "수행중요성": pm, "마감제외": "O",
        "재공품코드": (C.pick(codes, [169]) or [""])[0],
        "전기손익가산": info.get("전기손익가산", "X"), "당기손익포함": info.get("당기손익포함", "O"),
        "JS01_최소금액": int(pm * 0.05) // 1000000 * 1000000, "JS02_기간": 0, "JS02_기준금액": pm,
        "JS03_단위": 1000000, "JS03_최소금액": rnd(pm * 0.1, 1000000),
        "JS06_방향": "대변", "JS07_방향": "양쪽", "JS08_건수": 3, "JS11_일수": 30,
        "공휴일": [dt.date.fromisoformat(d) for d in hol],
        "키워드": list(KEYWORDS),
        "JS06_본계정": [c for c in sorted(codes, key=str) if C.is_revenue(c, names[c])],
        "JS06_허용": C.pick(codes, [101, 102, 103, 108, 110, 120, 255, 259]),
        "JS07_본계정": C.pick(codes, [114, 134, 179]), "JS07_허용": C.pick(codes, [101, 102, 103]),
        "JS09_대상": C.pick(codes, [108, 110, 120, 131, 251, 252, 253, 259]),
        "마감코드": C.pick(codes, [400]),
        "적용": {s: ("O" if (s not in ("JS10", "JS11") or has_wd) else "X") for s in SCN},
    }


def judge(lines, vouchers, tb_rows, P):
    """반환 {전표키: {마감, 원가, 본06, 본07, JS01.., 차변합계, 작성일}}."""
    grp = {r["코드"]: r["그룹"] for r in tb_rows}
    cnt = Counter(x["코드"] for x in lines)
    hol = set(P["공휴일"])
    kws = [k for k in P["키워드"] if k]
    use = P["적용"]
    out = {}

    def side_ok(x, d):
        return d == "양쪽" or (d == "대변" and x["대변"] != 0) or (d == "차변" and x["차변"] != 0)

    for v in vouchers:
        L = lines[v["s"]:v["e"] + 1]
        dr = sum(x["차변"] for x in L)
        wd = max((x["작성일"] for x in L if x["작성일"]), default=None)
        r = {"차변합계": dr, "대변합계": sum(x["대변"] for x in L), "작성일": wd}
        r["마감"] = int(any(x["코드"] in P["마감코드"] for x in L))
        r["원가"] = int(P["재공품코드"] != "" and any(x["코드"] == P["재공품코드"] for x in L)
                      and any(grp.get(x["코드"]) == "제조" and x["대변"] != 0 for x in L))
        amt = lambda x: abs(x["차변"]) + abs(x["대변"])
        d = v["일자"]
        raw = {
            "JS01": (d.weekday() >= 5 or d in hol) and dr >= P["JS01_최소금액"],
            "JS02": (P["기준일"] - dt.timedelta(days=P["JS02_기간"]) <= d <= P["기준일"]) and dr >= P["JS02_기준금액"],
            "JS03": any(amt(x) >= P["JS03_최소금액"] and round(amt(x)) % P["JS03_단위"] == 0 for x in L),
            "JS04": any(k.lower() in x["적요"].lower() for x in L for k in kws),
            "JS05": any(not x["적요"].strip() for x in L),
            "JS08": any(0 < cnt[x["코드"]] <= P["JS08_건수"] for x in L),
            "JS09": any(x["코드"] in P["JS09_대상"] and not x["거래처명"].strip() for x in L),
            "JS10": bool(wd and wd > P["기준일"]),
            "JS11": bool(wd and (wd - d).days >= P["JS11_일수"]),
        }
        for s in ("JS06", "JS07"):
            main, allow = P[s + "_본계정"], P[s + "_허용"]
            has_main = any(x["코드"] in main and side_ok(x, P[s + "_방향"]) for x in L)
            r["본" + s[2:]] = int(has_main)
            raw[s] = has_main and any(x["코드"] not in main and x["코드"] not in allow for x in L)
        excl = P["마감제외"] == "O" and r["마감"] == 1
        for s in SCN:
            r[s] = int(bool(raw[s]) and use[s] == "O" and not excl)
        out[v["키"]] = r
    return out
