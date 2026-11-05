# -*- coding: utf-8 -*-
r"""처음 한 번 - 거래처 루트를 정하고 설정 · VBA 파일을 만든다(Outlook 은 건드리지 않음).

  python -X utf8 setup.py --root "D:\거래처"  [--outlook-root 거래처] [--groups "01. 고객사,02. 협력사"]
  python -X utf8 setup.py                     이미 정한 루트로 VBA · 사용법 파일만 다시 만들기(스킬을 업데이트한 뒤)

만드는 것:
  %APPDATA%\outlook-client-mail\config.json     루트 위치(도구들이 읽음)
  <root>\_메일분류\mail_sort.json               사용자 설정(없을 때만 - 있으면 그대로 둠)
  <root>\_메일분류\vba\ClientSort.bas · ClientSortWatch.cls · ThisOutlookSession_붙여넣기.txt
                                                Outlook VBA 가져오기용(cp949 · CRLF, 경로가 채워져 있음)
  <root>\_메일분류\사용법.html                   VBA [사용법] 버튼이 여는 안내 페이지
  <root>\_메일분류\outlook_autostart.ps1        클래식 Outlook 자동 시작(바로가기용)
"""
import argparse
import datetime as dt
import glob
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from common import APP_CFG, CONF_DIR, DEFAULTS, write_json, write_text  # noqa: E402


def outlook_exe():
    pats = [r"C:\Program Files\Microsoft Office\root\Office*\OUTLOOK.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office*\OUTLOOK.EXE",
            r"C:\Program Files\Microsoft Office\Office*\OUTLOOK.EXE"]
    for p in pats:
        hit = sorted(glob.glob(p))
        if hit:
            return hit[-1]
    return r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root")
    ap.add_argument("--outlook-root")
    ap.add_argument("--groups", help="쉼표로 구분")
    a = ap.parse_args()
    if a.root:
        root = os.path.abspath(a.root)
        os.makedirs(root, exist_ok=True)
        if not os.environ.get("OCM_ROOT"):                              # 환경변수로 루트를 쓰는 중이면 기록 파일은 건드리지 않음
            write_json(APP_CFG, {"root": root}, 2)
    elif os.environ.get("OCM_ROOT"):
        root = os.path.abspath(os.environ["OCM_ROOT"])
    elif os.path.exists(APP_CFG):
        root = json.load(open(APP_CFG, encoding="utf-8"))["root"]
    else:
        raise SystemExit("처음에는 --root <거래처 루트 폴더> 를 주세요")
    conf = os.path.join(root, CONF_DIR)
    os.makedirs(conf, exist_ok=True)
    sp = os.path.join(conf, "mail_sort.json")
    if os.path.exists(sp):
        s = json.load(open(sp, encoding="utf-8"))
        print("설정 그대로 둠:", sp)
    else:
        s = dict(DEFAULTS, since=dt.date.today().isoformat())
        print("설정 만듦:", sp)
    if a.outlook_root:
        s["outlook_root"] = a.outlook_root
    if a.groups is not None:
        s["groups"] = [g.strip() for g in a.groups.split(",") if g.strip()]
    write_json(sp, s, 2)
    for g in s.get("groups") or []:
        os.makedirs(os.path.join(root, g), exist_ok=True)

    vb = os.path.join(conf, "vba")
    os.makedirs(vb, exist_ok=True)
    for name in ("ClientSort.bas", "ClientSortWatch.cls", "ThisOutlookSession_붙여넣기.txt"):
        text = open(os.path.join(SKILL, "vba", name), encoding="utf-8").read().replace("{{CONF_DIR}}", conf)
        with open(os.path.join(vb, name), "w", encoding="cp949", newline="\r\n") as f:     # 한글 Windows VBA 가져오기는 cp949
            f.write(text)
    ps = open(os.path.join(SKILL, "vba", "outlook_autostart.ps1"), encoding="utf-8").read()
    ps = ps.replace("{{CONF_DIR}}", conf).replace("{{OUTLOOK_EXE}}", outlook_exe())
    write_text(os.path.join(conf, "outlook_autostart.ps1"), ps, "utf-8-sig", "\r\n")
    html = open(os.path.join(SKILL, "references", "사용법.html"), encoding="utf-8").read()
    write_text(os.path.join(conf, "사용법.html"), html.replace("{{ROOT}}", root).replace("{{OUTLOOK_ROOT}}", s["outlook_root"]))
    print("거래처 루트:", root)
    print("Outlook 거래처 폴더:", s["outlook_root"], "(기본 사서함 최상위 아래)")
    print("VBA 파일:", vb)
    print("다음: references/install.md 순서대로 Outlook 에 VBA 를 가져오고, mail_sync_all.py --link-only 로 연결")


if __name__ == "__main__":
    main()
