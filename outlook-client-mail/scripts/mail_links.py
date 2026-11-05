# -*- coding: utf-8 -*-
r"""히스토리 원고의 {#번호} 를 원본 메일(.msg) 링크로 바꾼다.

  python -X utf8 mail_links.py <거래처> <원고.md> [<출력.md>]     출력을 빼면 거래처의 메일 히스토리 문서에 쓴다

  {#29} → [#29](<메일/2607/260729_1330_보냄_….msg>)  - 출력 문서 위치 기준 상대 경로(VS Code · Obsidian 등에서 누르면 Outlook 으로 열림).
  이미 링크인 [#n](...) 은 그대로 둔다. 색인에 없는 번호는 '#n(원본 없음)'.
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import history_path, resolve, write_text  # noqa: E402
from mail_search import load  # noqa: E402


def main():
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    d, src = resolve(sys.argv[1]), sys.argv[2]
    dst = sys.argv[3] if len(sys.argv) > 3 else history_path(d)
    recs = load(d)
    base = os.path.dirname(os.path.abspath(dst))
    text = open(src, encoding="utf-8").read()

    def rep(m):
        n = int(m.group(1))
        if not 1 <= n <= len(recs) or not recs[n - 1].get("msg"):
            return "#%d(원본 없음)" % n
        rel = os.path.relpath(os.path.join(d, recs[n - 1]["msg"]), base).replace("\\", "/")
        return "[#%d](<%s>)" % (n, rel)
    out = re.sub(r"\{#(\d+)\}", rep, text)
    write_text(dst, out)
    print("링크 %d개 → %s" % (len(re.findall(r"\{#\d+\}", text)), dst))


if __name__ == "__main__":
    main()
