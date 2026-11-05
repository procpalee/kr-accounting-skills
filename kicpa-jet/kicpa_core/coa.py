# -*- coding: utf-8 -*-
"""계정 분류 - 더존 3자리 계정 체계(101 현금 … 401 매출 … 801 판관비)를 기준으로 계정의 성격을 정한다.

ERP 코드가 더존 코드에 자릿수를 붙인 형태(1020000 = 102, 1030001 = 103)면 앞 3자리로 본다.
TB 에 구분(자산 · 부채 · 자본 · 손익) 열이 있으면 그 값이 우선이다. 코드로도 구분으로도 못 정하면 '미분류'.
"""
import re


def cls3(code):
    """계정코드 → 더존 3자리 분류 번호(없으면 None)."""
    if isinstance(code, bool):
        return None
    if isinstance(code, int):
        s = str(code)
    else:
        s = str(code or "").strip()
        if not s.isdigit():
            return None
    if len(s) < 3:
        return None
    return int(s[:3])


def kind_by_class(c):
    if c is None:
        return None
    if 101 <= c <= 250:
        return "자산"
    if 251 <= c <= 330:
        return "부채"
    if 331 <= c <= 399:
        return "자본"
    if 500 <= c <= 799:
        return "제조"
    if 400 <= c <= 999:
        return "손익"
    return None


GIVEN = {"자산": "자산", "부채": "부채", "자본": "자본", "손익": "손익", "수익": "손익", "비용": "손익",
         "제조": "제조", "제조원가": "제조", "원가": "제조"}


def kind_of(code, name="", given=None):
    """자산 · 부채 · 자본 · 손익 · 제조 · 미분류."""
    c = cls3(code)
    k = kind_by_class(c)
    g = GIVEN.get(str(given or "").strip())
    if g == "손익" and k == "제조":
        return "제조"
    if g:
        return g
    return k or "미분류"


RE_NAME = re.compile(r"이익잉여금|이익준비금|적립금|결손금|이월이익|처분전이익")


def is_retained(code, name, kind):
    c = cls3(code)
    if kind != "자본":
        return False
    if c is not None and 351 <= c <= 380:
        return True
    return bool(RE_NAME.search(str(name or "")))


def jet_group(code, name, kind):
    """분개테스트 완전성 그룹: 재무상태표 · 이익잉여금 · 손익 · 제조 · 미분류."""
    if kind in ("자산", "부채", "자본"):
        return "이익잉여금" if is_retained(code, name, kind) else "재무상태표"
    if kind in ("손익", "제조"):
        return kind
    return "미분류"


def pick(codes, classes):
    """codes(분개장에 나온 계정) 중 더존 분류 번호가 classes 안인 것."""
    cs = set(classes)
    return sorted((c for c in codes if cls3(c) in cs), key=str)


def is_revenue(code, name):
    c = cls3(code)
    n = str(name or "")
    return c is not None and 401 <= c <= 430 and "매출" in n and "원가" not in n
