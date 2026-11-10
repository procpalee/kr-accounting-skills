# -*- coding: utf-8 -*-
"""xlsync_harvest - 쌓인 비교 기록(_xlsync/**/diff_*.json)에서 사람이 반복해서 고치는 것을 집계한다.

    python -X utf8 xlsync_harvest.py <폴더> [--out 결과.json] [--min-files 3]

읽기만 한다. 출력에는 서식 값(글꼴·색·표시 형식·너비)과 글의 변화 종류만 나오고 칸의 글·숫자·파일 이름·시트 이름은 나오지 않는다.
한 파일에서 여러 번 고친 것은 파일 1개로 센다 - 반복 여부는 '파일 수'로 본다.
"""
import argparse
import collections
import glob
import hashlib
import json
import os
import re
import sys

PUNCT = re.compile(r"[\s\.,·ㆍ\-–—_:;()\[\]{}'\"“”‘’/~]+")


def _is_formula(v):
    return isinstance(v, dict) and "f" in v


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _blank(v):
    return v in (None, "")


def text_change_kind(a, b):
    """글이 어떻게 바뀌었는지 종류만 돌려준다."""
    if a.strip() == b.strip() and a != b:
        return "앞뒤 공백만"
    if re.sub(r"\s+", "", a) == re.sub(r"\s+", "", b):
        return "띄어쓰기만"
    if PUNCT.sub("", a) == PUNCT.sub("", b):
        for old, new, name in (("—", "-", "줄표를 하이픈으로"), (" · ", "·", "가운뎃점 앞뒤 공백 제거"), (" ·", "·", "가운뎃점 앞뒤 공백 제거"),
                               ("· ", "·", "가운뎃점 앞뒤 공백 제거"), ("ㆍ", "·", "가운뎃점 문자 통일")):
            if old in a and new in b and a.replace(old, new) == b:
                return name
        if a.rstrip(".") == b.rstrip("."):
            return "끝 마침표"
        return "문장부호만"
    if b.startswith(a):
        return "뒤에 글 덧붙임"
    if a.startswith(b):
        return "뒤 글 삭제"
    if a in b:
        return "앞뒤에 글 덧붙임"
    return "내용 고침"


def value_kind(a, b):
    if _is_formula(a) and _is_formula(b):
        return "수식을 다른 수식으로"
    if _is_formula(a):
        return "수식을 값으로" if not _blank(b) else "수식 삭제"
    if _is_formula(b):
        return "값을 수식으로" if not _blank(a) else "빈 칸에 수식"
    if _blank(a) and _is_num(b):
        return "빈 칸에 숫자"
    if _blank(a):
        return "빈 칸에 글"
    if _blank(b):
        return "숫자 삭제" if _is_num(a) else "글 삭제"
    if _is_num(a) and _is_num(b):
        return "숫자를 다른 숫자로"
    if isinstance(a, str) and isinstance(b, str):
        return "글: " + text_change_kind(a, b)
    return "형식이 바뀜"


def _cells(c):
    return int(c.get("cells") or c.get("count") or 1)


def harvest(root):
    files = sorted(glob.glob(os.path.join(root, "**", "_xlsync", "**", "diff_*.json"), recursive=True))
    per = collections.defaultdict(lambda: collections.defaultdict(set))
    hits = collections.Counter()
    total, used, ids = 0, 0, set()
    for fp in files:
        total += 1
        try:
            with open(fp, encoding="utf-8") as f:
                res = json.load(f)["result"]
        except (OSError, ValueError, KeyError):
            continue
        used += 1
        doc = hashlib.sha1(os.path.dirname(fp).lower().encode("utf-8")).hexdigest()[:10]
        ids.add(doc)
        for c in res.get("changes", []):
            k = c.get("kind")
            if k == "style":
                key = ("서식", str(c.get("attr")), json.dumps(c.get("before"), ensure_ascii=False, sort_keys=True), json.dumps(c.get("after"), ensure_ascii=False, sort_keys=True))
            elif k == "value":
                key = ("값", value_kind(c.get("before"), c.get("after")), "", "")
            elif k in ("width", "height"):
                b, a = c.get("before"), c.get("after")
                if c.get("attr") == "hidden":
                    key = ("열행", "%s 숨김" % k, str(b), str(a))
                else:
                    move = "늘림" if _is_num(a) and _is_num(b) and a > b else "줄임" if _is_num(a) and _is_num(b) and a < b else "지정/해제"
                    key = ("열행", k, move, "")
            else:
                key = ("구조", str(k), str(c.get("attr") or ""), "")
            per[key]["doc"].add(doc)
            hits[key] += 1
    return dict(total=total, used=used, docs=len(ids)), [
        dict(group=k[0], what=k[1], before=k[2], after=k[3], files=len(v["doc"]), hits=hits[k]) for k, v in per.items()]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("--out")
    ap.add_argument("--min-files", type=int, default=3)
    a = ap.parse_args()
    stat, rows = harvest(a.root)
    rows.sort(key=lambda r: (-r["files"], -r["hits"]))
    print("비교 기록 %(total)d개 (읽음 %(used)d) · 문서 %(docs)d개" % stat)
    print("파일수\t건수\t구분\t무엇\t전\t후")
    for r in rows:
        if r["files"] >= a.min_files:
            print("%d\t%d\t%s\t%s\t%s\t%s" % (r["files"], r["hits"], r["group"], r["what"], r["before"], r["after"]))
    if a.out:
        with open(a.out, "w", encoding="utf-8") as f:
            json.dump(dict(stat=stat, rows=rows), f, ensure_ascii=False, indent=1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
