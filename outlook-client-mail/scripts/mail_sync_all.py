# -*- coding: utf-8 -*-
r"""모든 거래처 한 번에 - 폴더 연결 맞추기 · 거래처 목록 · 수집 · 요약.

  python -X utf8 mail_sync_all.py              연결 맞추기 + 거래처 목록 + 모든 거래처 수집 + 요약   ← Outlook VBA 가 켤 때 부름
  python -X utf8 mail_sync_all.py --link-only  연결 맞추기 + 거래처 목록만                           ← VBA [폴더 연결 맞추기] 버튼
  python -X utf8 mail_sync_all.py --status     수집 없이 요약만                                       ← Claude 가 히스토리 정리 전에
  python -X utf8 mail_sync_all.py --mark <거래처> <번호>   그 거래처 히스토리를 #번호까지 반영했다고 기록
  결과는 <root>\_메일분류\last_sync.txt 에도 남긴다(--status 제외).

요약 줄: <거래처>  전체 N통 · 새로 모음 k통 · 히스토리 반영 #d 까지 → 정리할 메일 #d+1 ~ #N
  그 아래: 새 첨부(반영 뒤 받은 메일의 첨부 폴더) · 대용량 링크(만료일 · 다운로드 폴더에 받았는지)
  · 역할 모르는 보낸 사람(거래처 · 외부 · 우리 어디에도 없는 주소 - 개인 메일을 쓰는 거래처 담당자 등)
"""
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from common import clients, conf_dir, load_mail_cfg, meta, resolve, write_json, write_text  # noqa: E402

BIG_RE = re.compile(r"대용량|bigfile|다운로드 기한|다운로드 기간", re.I)
BIG_FILE_RE = re.compile(r"([^\s<>|][^<>\n|]*?\.(?:zip|7z|rar|xlsx?|xlsm|pdf|hwpx?|docx?|pptx?|csv))\s*<https?://", re.I)
BIG_DUE_RE = re.compile(r"(?:기한|기간|만료)[^\d\n]{0,10}(\d{4}[-./]\d{1,2}[-./]\d{1,2})")
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")


def active_clients():
    """메일 설정이 있고 Outlook 에서 보관하지 않은 거래처."""
    out = []
    for c in clients():
        cfg = load_mail_cfg(c["dir"])
        if cfg is not None and cfg.get("outlook", {}).get("active", True) is not False:
            out.append(c["dir"])
    return out


def read_state(d):
    p = meta(d, "mail_state.json")
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else {}


def index(d):
    p = meta(d, "mail_index.jsonl")
    return [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()] if os.path.exists(p) else []


def notes(idx, after):
    """반영 뒤 받은 메일 중 첨부 · 대용량 링크 → [(번호, 설명)]."""
    out = []
    for n, r in enumerate(idx, 1):
        if n <= after or r.get("dir") != "받음" or r.get("from", {}).get("role") in ("나", "우리"):
            continue
        bits = []
        if r.get("attach_dir"):
            bits.append("첨부 %s" % r["attach_dir"].split("\\")[-1][:50])
        body = r.get("body", "")
        if BIG_RE.search(body):
            names = list(dict.fromkeys(m.strip() for m in BIG_FILE_RE.findall(body)))
            due = BIG_DUE_RE.search(body)
            have = [x for x in names if os.path.exists(os.path.join(DOWNLOADS, x))]
            bits.append("대용량 링크%s%s%s" % (" " + ", ".join(names[:3]) if names else "",
                                             " · 만료 %s" % due.group(1) if due else "",
                                             " · 다운로드 폴더에 있음" if have else (" · 아직 안 받음" if names else "")))
        if bits:
            out.append((n, "%s %s · %s" % (r["at"][5:10], r.get("subject", "")[:40], " / ".join(bits))))
    return out


def unknown_senders(idx):
    """역할을 모르는 보낸 사람 → [(주소, 이름, 통수)] (많은 순)."""
    seen = {}
    for r in idx:
        f = r.get("from", {})
        if r.get("dir") == "받음" and not f.get("role") and f.get("addr"):
            ad = f["addr"].lower()
            nm, k = seen.get(ad, (f.get("name", ""), 0))
            seen[ad] = (nm, k + 1)
    return sorted(((ad, nm, k) for ad, (nm, k) in seen.items()), key=lambda x: -x[2])


def link_report(lines):
    """[폴더 연결 맞추기] 결과 - 상태 요약 + 바뀐 것 + 확인할 것."""
    on, off, none = [], [], []
    for c in clients():
        link = (load_mail_cfg(c["dir"]) or {}).get("outlook", {})
        name = os.path.basename(c["dir"])
        if not link.get("entry_id"):
            none.append(name)
        elif link.get("active") is False:
            off.append("%s → %s" % (name, link.get("path", "")))
        else:
            on.append("%s → %s" % (name, link.get("path", "")))
    made = [m for m in lines if m.startswith(("만들기", "연결 -", "메일 설정 만듦", "보관"))]
    warn = [m for m in lines if m.startswith(("연결 끊김", "확인 필요")) or "실패" in m]
    loose = [m for m in lines if m.startswith("연결 안 된 Outlook 폴더")]
    out = ["자동 분류 중인 거래처 %d곳 · Outlook 에서 보관한 거래처 %d곳 · 연결 안 된 거래처 %d곳" % (len(on), len(off), len(none)), ""]
    out += (["[이번에 새로 연결 · 만든 것]"] + ["  " + m for m in made] + [""]) if made else ["[이번에 바뀐 것] 없음", ""]
    if warn:
        out += ["[확인 필요]"] + ["  " + m for m in warn] + [""]
    if off:
        out += ["[Outlook 에서 보관 - 자동 분류 안 함]"] + ["  " + m for m in off] + [""]
    if loose:
        out += ["[어느 거래처와도 연결 안 된 Outlook 폴더 - 그대로 둠]"] + ["  " + m.split(" - ", 1)[-1].split("(")[0] for m in loose] + [""]
    out += ["[자동 분류 중인 거래처]"] + ["  " + m for m in on]
    return out


def save_result(lines):
    import datetime as dt
    write_text(os.path.join(conf_dir(), "last_sync.txt"), "%s\n%s\n" % (dt.datetime.now().isoformat(timespec="seconds"), "\n".join(lines)), "utf-8-sig")
    print("\n".join(lines))


def main():
    a = sys.argv[1:]
    if a[:1] == ["--mark"]:
        d = resolve(a[1])
        st = read_state(d)
        st["digested_upto"] = int(a[2])
        write_json(meta(d, "mail_state.json"), st)
        print("%s 히스토리 반영 #%s 까지로 기록" % (os.path.basename(d), a[2]))
        return
    lines = []
    if "--status" not in a:
        try:
            from outlook_folders import run as link_run
            lines += link_run(True)
        except Exception as ex:  # noqa: BLE001
            lines.append("Outlook 폴더 연결 맞추기 실패: %s" % ex)
        try:
            from client_map import build
            build()
        except Exception as ex:  # noqa: BLE001
            lines.append("거래처 목록(Client_Map.tsv) 갱신 실패: %s" % ex)
    if "--link-only" in a:
        save_result(link_report(lines))
        return
    out = ["[Outlook 폴더] " + m for m in lines]
    for d in active_clients():
        before = len(index(d))
        if "--status" not in a:
            r = subprocess.run([sys.executable, "-X", "utf8", os.path.join(HERE, "mail_sync.py"), d],
                               capture_output=True, text=True, encoding="utf-8", errors="replace")
            if r.returncode:
                out.append("%s  수집 실패: %s" % (os.path.basename(d), (r.stderr or r.stdout).strip()[-300:]))
                continue
        idx, st = index(d), read_state(d)
        n, done = len(idx), st.get("digested_upto", 0)
        todo = "정리할 메일 #%d ~ #%d" % (done + 1, n) if n > done else "정리할 메일 없음"
        out.append("%s  전체 %d통 · 새로 모음 %d통 · 히스토리 반영 #%d 까지 → %s" % (os.path.basename(d), n, n - before, done, todo))
        for no, desc in notes(idx, done):
            out.append("    #%d  %s" % (no, desc))
        unknown = unknown_senders(idx)
        if unknown:
            out.append("    역할 모르는 보낸 사람 %d명  %s" % (len(unknown), ", ".join("%s %s(%d통)" % (nm, ad, k) for ad, nm, k in unknown[:5])
                                                        + (" …" if len(unknown) > 5 else "")))
    if "--status" not in a:
        save_result(out)
    else:
        print("\n".join(out) or "메일 설정(_mail\\mail.json)이 있는 거래처가 없습니다")


if __name__ == "__main__":
    main()
