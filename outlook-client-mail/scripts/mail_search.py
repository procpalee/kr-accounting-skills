# -*- coding: utf-8 -*-
r"""메일 찾기 · 열기 - mail_sync 가 모은 한 거래처의 메일 색인에서 찾는다(Outlook 을 열지 않음).

  python -X utf8 mail_search.py <거래처> [키워드 ...] [--from 이름|주소] [--since YYYY-MM-DD] [--until YYYY-MM-DD]
                                        [--dir 받음|보냄] [--attach] [--after N] [--body] [--limit 30] [--open 번호]

  키워드는 모두 들어간 메일만(제목 · 본문 · 보낸/받는 사람 · 첨부 이름에서). 번호(#)는 색인 줄 번호 = 메일 히스토리의 #번호.
  --after N : N 번 다음 메일만(새로 모은 메일 정리용) · --body : 본문과 히스토리용 링크까지 출력
  --open 번호 : 그 메일 원본(.msg)을 Outlook 으로 연다.
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import history_path, meta, resolve  # noqa: E402


def load(d):
    p = meta(d, "mail_index.jsonl")
    if not os.path.exists(p):
        raise SystemExit("메일 색인이 없습니다 - 먼저 mail_sync.py 로 모으세요: " + p)
    return [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]


def link_of(d, msg):
    """메일 히스토리 문서에서 원본까지의 상대 경로."""
    return os.path.relpath(os.path.join(d, msg), os.path.dirname(history_path(d))).replace("\\", "/")


def text_of(r):
    ppl = [r["from"]] + r["to"] + r["cc"]
    return " ".join([r["subject"], r["body"], " ".join(r["attachments"])] + ["%s %s" % (x["name"], x["addr"]) for x in ppl]).lower()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("client")
    ap.add_argument("kw", nargs="*")
    ap.add_argument("--from", dest="frm")
    ap.add_argument("--since")
    ap.add_argument("--until")
    ap.add_argument("--dir")
    ap.add_argument("--attach", action="store_true", help="첨부 있는 메일만")
    ap.add_argument("--limit", type=int, default=30)
    ap.add_argument("--open", type=int)
    ap.add_argument("--body", action="store_true", help="본문까지 출력(히스토리 정리용)")
    ap.add_argument("--after", type=int, default=0, help="이 번호 다음 메일만")
    a = ap.parse_args()
    d = resolve(a.client)
    recs = load(d)
    if a.open:
        r = recs[a.open - 1]
        print("#%d %s %s" % (a.open, r["at"], r["subject"]))
        os.startfile(os.path.join(d, r["msg"]))                         # noqa: S606 - Outlook 으로 열기
        return
    hits = []
    for i, r in enumerate(recs, 1):
        if i <= a.after:
            continue
        t = text_of(r)
        if any(k.lower() not in t for k in a.kw):
            continue
        if a.frm and a.frm.lower() not in ("%s %s" % (r["from"]["name"], r["from"]["addr"])).lower():
            continue
        if a.since and r["at"] < a.since:
            continue
        if a.until and r["at"][:10] > a.until:
            continue
        if a.dir and r["dir"] != a.dir:
            continue
        if a.attach and not r["attachments"]:
            continue
        hits.append((i, r))
    print("%s - %d통 중 %d통" % (os.path.basename(d), len(recs), len(hits)))
    for i, r in hits[-a.limit:]:
        who = r["from"]["name"] if r["dir"] == "받음" else "→ " + ", ".join(x["name"] or x["addr"] for x in r["to"][:2])
        att = " [첨부: %s]" % ", ".join(r["attachments"][:3]) if r["attachments"] else ""
        print("#%-4d %s %s %-14s %s%s" % (i, r["at"].replace("T", " "), r["dir"], who[:14], r["subject"][:60], att))
        if a.body:
            cc = ", ".join(x["name"] or x["addr"] for x in r["cc"])
            print("      받는 사람: %s%s" % (", ".join(x["name"] or x["addr"] for x in r["to"]), (" / 참조: " + cc) if cc else ""))
            print("      링크: [#%d](<%s>)" % (i, link_of(d, r["msg"])))
            print("      " + r["body"][:3000].replace("\n", "\n      ") + "\n")
    if len(hits) > a.limit:
        print("... 앞의 %d통 생략(--limit)" % (len(hits) - a.limit))
    print("원본 열기: mail_search.py %s --open <번호>" % a.client)


if __name__ == "__main__":
    main()
