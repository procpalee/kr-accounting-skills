# -*- coding: utf-8 -*-
r"""Outlook VBA(ClientSort)가 읽는 거래처 목록(Client_Map.tsv)을 만든다 - 자동 분류의 단일 기준.

  python -X utf8 client_map.py            다시 만들기(+ VBA 가 '기억'한 주소를 거래처 mail.json 에 반영)
  python -X utf8 client_map.py --show     만들 내용만 보기

입력: <root>\[그룹\]<거래처>\_mail\mail.json
      <root>\_메일분류\mail_sort.json        감시 폴더 · outlook_root · 단체 메일 기준(사용자 설정)
      <root>\_메일분류\Client_Map_추가.tsv   VBA [거래처로 분류] 에서 '기억'한 주소(거래처명 \t 주소 \t 값 \t 날짜)
출력: <root>\_메일분류\Client_Map.tsv (UTF-8 BOM, CRLF)
  #client_root \t <outlook_root>
  #watch \t <감시 폴더>                (받은 · 보낸 편지함은 VBA 가 항상 감시)
  #max_recipients \t 15
  #exclude \t <제목에 이 말이 있으면 분류 안 함>
  #python · #mail_sync · #mail_all · #help \t 경로   (VBA 가 창 없이 부르는 도구)
  거래처명 \t Outlook 폴더(outlook_root 기준) \t 제목 키워드(;) \t 주소 · 도메인(;) \t 담당자 주소(;) \t 거래처 폴더 \t Outlook 폴더 ID
"""
import datetime as dt
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import clients, conf_dir, load_mail_cfg, mail_cfg_path, settings, write_text, write_json  # noqa: E402

SKILL = os.path.dirname(HERE)
PYW = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")      # VBA 가 창 없이 부를 때
if not os.path.exists(PYW):
    PYW = sys.executable


def extras():
    p = os.path.join(conf_dir(), "Client_Map_추가.tsv")
    out = []
    if os.path.exists(p):
        for line in open(p, encoding="utf-8-sig"):
            parts = line.rstrip("\r\n").split("\t")
            if len(parts) >= 3 and not line.startswith("#"):
                out.append((parts[0].strip(), parts[1].strip(), parts[2].strip().lower()))
    return out


def build(write=True):
    s, ex, rows = settings(), extras(), []
    for c in clients():
        name, mj = c["name"], load_mail_cfg(c["dir"])
        learned = [v for n, k, v in ex if n == name and k == "주소"]
        if mj is not None and learned:                                  # VBA 가 기억한 주소 → mail.json(수집 · 역할 판단도 같이 씀)
            have = {x.lower() for x in mj["company"].get("addresses", []) + mj["company"].get("domains", [])}
            new = [v for v in learned if v not in have]
            if new and write:
                mj["company"]["addresses"] = mj["company"].get("addresses", []) + new
                write_json(mail_cfg_path(c["dir"]), mj, 2)
        comp, oth = (mj or {}).get("company", {}), (mj or {}).get("others", {})
        addrs = [x.strip().lower() for x in comp.get("addresses", []) + comp.get("domains", []) + oth.get("addresses", [])
                 + oth.get("domains", []) + learned if x.strip()]
        contacts = [x.strip().lower() for x in comp.get("addresses", []) + learned if "@" in x]
        kws = [name] + (mj or {}).get("keywords", []) + [v for n, k, v in ex if n == name and k == "키워드"]
        link = (mj or {}).get("outlook", {})
        if link.get("active") is False:                                 # Outlook 에서 보관한 거래처 - 자동 분류 · 저장 안 함
            continue
        default_path = (c["group"] + "\\" if c["group"] else "") + os.path.basename(c["dir"])
        rows.append([name, link.get("path") or default_path, ";".join(dict.fromkeys(k for k in kws if k)),
                     ";".join(dict.fromkeys(addrs)), ";".join(dict.fromkeys(contacts)),
                     c["dir"] if mj is not None else "", link.get("entry_id", "")])
    lines = ["#생성\t%s\tclient_map.py - 직접 고치지 말고 mail.json · mail_sort.json · Client_Map_추가.tsv 를 고친다"
             % dt.datetime.now().isoformat(timespec="seconds"),
             "#client_root\t%s" % s["outlook_root"], "#python\t%s" % PYW,
             "#mail_sync\t%s" % os.path.join(HERE, "mail_sync.py"), "#mail_all\t%s" % os.path.join(HERE, "mail_sync_all.py"),
             "#help\t%s" % os.path.join(conf_dir(), "사용법.html")] + ["#watch\t%s" % w for w in s["watch"]] + \
            ["#max_recipients\t%d" % int(s["max_recipients"])] + ["#exclude\t%s" % x for x in s["exclude_subjects"]] + \
            ["#거래처명\tOutlook 폴더\t제목키워드\t주소\t담당자\t거래처 폴더(설정이 있을 때 - 분류 직후 수집)\tOutlook 폴더 ID"] + \
            ["\t".join(r) for r in rows]
    if write:
        p = os.path.join(conf_dir(), "Client_Map.tsv")
        old = open(p, encoding="utf-8-sig").read().replace("\r\n", "\n").split("\n", 1)[-1] if os.path.exists(p) else None
        body = "\n".join(lines[1:]) + "\n"
        if old != body:                                                 # 내용이 같으면 그대로(VBA 가 다시 읽지 않게)
            write_text(p, lines[0] + "\n" + body, "utf-8-sig", "\r\n")
    return lines


def main():
    lines = build(write="--show" not in sys.argv)
    print("\n".join(lines))
    if "--show" not in sys.argv:
        print("→ %s" % os.path.join(conf_dir(), "Client_Map.tsv"))


if __name__ == "__main__":
    main()
