# -*- coding: utf-8 -*-
r"""메일 수집 - 한 거래처의 Outlook 폴더에 있는 메일(받음 · 보냄)을 거래처 폴더로 저장한다. Outlook 은 읽기만 한다.

  python -X utf8 mail_sync.py <거래처 폴더 | 이름> [--dry-run] [--since YYYY-MM-DD] [--full]

거래처 메일 = 그 거래처 Outlook 폴더(_mail\mail.json 의 outlook.entry_id 로 연결)에 있는 메일.
VBA ClientSort 가 넣었든 사람이 끌어다 넣었든 같다. 보낸 메일은 VBA 가 넣는 사본으로 들어온다.
VBA 가 그 폴더에 메일이 들어오는 순간 이 도구를 부른다. 늦게 넣은 메일도 since 부터 매번 훑어 모은다(이미 모은 메일은 건너뜀).

결과:
  <거래처>\메일\<YYMM>\<YYMMDD_HHMM>_<받음|보냄>_<상대>_<제목>.msg   원본(더블클릭하면 Outlook 으로 열림)
  <거래처>\첨부\<YYMMDD_HHMM>_#<번호>_<보낸 사람>_<제목>\         받은 메일 첨부(서명 이미지 · 일정 · 연락처 제외)
  <거래처>\_mail\mail_index.jsonl                                  메일 한 통 = 한 줄
  <거래처>\_mail\mail_state.json                                   마지막 수집 시각
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import (attach_dir, bare, load_mail_cfg, meta, resolve, save_dir, settings, write_json)  # noqa: E402

PR_MSGID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
OL_MSG_UNICODE = 9
ATT_SKIP = re.compile(r"^(image\d+\.(png|jpe?g|gif|bmp)|.+\.(ics|vcf))$", re.I)     # 서명 이미지 · 일정 · 연락처
QUOTE_RE = re.compile(r"^\s*(-{2,}\s*(Original Message|원본 메시지|Forwarded message)|From:\s|보낸 사람:\s|발신:\s|On .+ wrote:|-----\s*$)", re.I | re.M)


def load_cfg(d):
    c = load_mail_cfg(d)
    if c is None:
        raise SystemExit("거래처 설정이 없습니다: %s - mail_sync_all.py --link-only 를 먼저 돌리세요" % d)
    s = settings()
    low = lambda xs: {x.lower().strip() for x in xs or [] if x.strip()}  # noqa: E731
    c["_company"] = low(c.get("company", {}).get("addresses"))
    c["_company_dom"] = low(c.get("company", {}).get("domains"))
    c["_other"] = low(c.get("others", {}).get("addresses"))
    c["_other_dom"] = low(c.get("others", {}).get("domains"))
    c["_team"] = low(s.get("team"))
    return c


def smtp_of_entry(ae):
    try:
        if ae.Type == "EX":
            u = ae.GetExchangeUser()
            if u:
                return (u.PrimarySmtpAddress or "").lower()
        return (ae.Address or "").lower()
    except Exception:  # noqa: BLE001
        return ""


def people(m):
    try:
        s = smtp_of_entry(m.Sender) if m.Sender else (m.SenderEmailAddress or "").lower()
    except Exception:  # noqa: BLE001
        s = (getattr(m, "SenderEmailAddress", "") or "").lower()
    to, cc = [], []
    for r in m.Recipients:
        a = smtp_of_entry(r.AddressEntry) if r.AddressEntry else (r.Address or "").lower()
        (cc if r.Type == 2 else to).append((r.Name, a))
    return (getattr(m, "SenderName", ""), s), to, cc


def role(addr, c, me):
    """나 · 거래처 · 외부 · 우리(내 도메인이거나 team 목록) · ''(모름)."""
    d = addr.split("@")[-1] if "@" in addr else ""
    if addr == me:
        return "나"
    if addr in c["_company"] or d in c["_company_dom"]:
        return "거래처"
    if addr in c["_other"] or d in c["_other_dom"]:
        return "외부"
    if addr in c["_team"] or (d and "@" in me and d == me.split("@")[-1]):
        return "우리"
    return ""


def clean_body(b):
    """인용된 이전 메일(Original Message · 보낸 사람: …) 앞까지만, 6000자까지."""
    b = (b or "").replace("\r\n", "\n")
    mm = QUOTE_RE.search(b)
    head = b[:mm.start()] if mm and mm.start() > 20 else b
    head = re.sub(r"\n{3,}", "\n\n", head).strip()
    return head[:6000], len(b)


def safe(s, n):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]', " ", s or "").strip()
    s = re.sub(r"\s+", " ", s)
    return s[:n].strip() or "제목없음"


def msg_id(m):
    try:
        v = m.PropertyAccessor.GetProperty(PR_MSGID)
        if v:
            return v
    except Exception:  # noqa: BLE001
        pass
    raw = "%s|%s|%s" % (m.SentOn, m.Subject, getattr(m, "SenderEmailAddress", ""))
    return "h:" + hashlib.md5(raw.encode("utf-8", "replace")).hexdigest()


def save_attachments(m, dst):
    """첨부(숨김 · 서명 이미지 · 일정 제외)를 받은 이름 그대로 저장 → [이름]. 하나도 없으면 폴더를 만들지 않는다."""
    out = []
    for at in m.Attachments:
        try:
            if getattr(at, "Type", 1) != 1 or ATT_SKIP.match(at.FileName or ""):
                continue
            try:
                if at.PropertyAccessor.GetProperty(PR_HIDDEN):
                    continue
            except Exception:  # noqa: BLE001 - 속성이 없으면 보이는 첨부
                pass
            os.makedirs(dst, exist_ok=True)
            name = re.sub(r'[\\/:*?"<>|]', "_", at.FileName)
            p, k = os.path.join(dst, name), 2
            while os.path.exists(p):
                p = os.path.join(dst, "%s (%d)%s" % (os.path.splitext(name)[0], k, os.path.splitext(name)[1]))
                k += 1
            at.SaveAsFile(p)
            out.append(os.path.basename(p))
        except Exception as ex:  # noqa: BLE001
            print("첨부 저장 실패:", getattr(at, "FileName", "?"), ex)
    return out


def client_folder(ns, c, d):
    """연결된 Outlook 폴더(entry_id). 연결이 없으면 outlook_root 아래에서 이름이 같은 폴더를 찾는다."""
    link = c.get("outlook", {})
    if link.get("entry_id"):
        try:
            return ns.GetFolderFromID(link["entry_id"], link.get("store_id") or ns.DefaultStore.StoreID)
        except Exception:  # noqa: BLE001 - 끊긴 연결은 outlook_folders.py 가 보고
            pass
    name = bare(os.path.basename(d))
    try:
        top = ns.DefaultStore.GetRootFolder().Folders.Item(settings()["outlook_root"])
    except Exception:  # noqa: BLE001
        return None
    stack = [top]
    while stack:
        f = stack.pop()
        for sf in f.Folders:
            if bare(sf.Name) == name:
                return sf
            stack.append(sf)
    return None


def main():
    """잠금: VBA 가 메일마다 이 도구를 부르므로 같은 거래처가 겹치면 뒤엣것은 '다시 돌기' 표시만 남기고 끝낸다."""
    ap = argparse.ArgumentParser()
    ap.add_argument("client")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--since")
    ap.add_argument("--full", action="store_true", help="수집 시각 무시하고 since 부터 다시 훑기(이미 있는 메일은 건너뜀)")
    a = ap.parse_args()
    d = resolve(a.client)
    if a.dry_run:
        return run(a, d)
    os.makedirs(meta(d), exist_ok=True)
    lock, again = meta(d, "mail_sync.lock"), meta(d, "mail_sync.rerun")
    if os.path.exists(lock) and dt.datetime.now().timestamp() - os.path.getmtime(lock) > 900:
        os.remove(lock)                                                 # 15분 넘은 잠금 = 죽은 실행
    try:
        os.close(os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY))
    except FileExistsError:
        open(again, "w").close()
        print("다른 수집이 도는 중 - 끝나면 한 번 더 돈다")
        return
    try:
        while True:
            if os.path.exists(again):
                os.remove(again)
            run(a, d)
            if not os.path.exists(again):
                break
    finally:
        os.remove(lock)


def run(a, d):
    c = load_cfg(d)
    idx_p, st_p = meta(d, "mail_index.jsonl"), meta(d, "mail_state.json")
    have, n0 = set(), 0
    if os.path.exists(idx_p):
        for line in open(idx_p, encoding="utf-8"):
            if not line.strip():
                continue
            n0 += 1
            try:
                have.add(json.loads(line).get("id"))
            except ValueError:
                pass
    state = json.load(open(st_p, encoding="utf-8")) if os.path.exists(st_p) else {}
    base = a.since or c.get("since") or settings()["since"]
    since_dt = dt.datetime.fromisoformat(base[:19])
    out_dir = save_dir(d)

    import win32com.client
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    me = (ns.CurrentUser.AddressEntry and smtp_of_entry(ns.CurrentUser.AddressEntry)) or ""
    f = client_folder(ns, c, d)
    if f is None:
        print("%s - Outlook 거래처 폴더 없음(mail_sync_all.py --link-only 로 연결)" % os.path.basename(d))
        return
    flt = "[ReceivedTime] >= '%s'" % since_dt.strftime("%m/%d/%Y %H:%M")
    new, seen_now, newest = [], set(), state.get("last_received", "")

    def walk(fo, rel):
        nonlocal newest
        if fo.DefaultItemType == 0:
            try:
                items = fo.Items.Restrict(flt)
            except Exception:  # noqa: BLE001
                items = []
            for m in items:
                if getattr(m, "Class", 0) != 43:                       # 메일만(회의 요청 · 보고서 제외)
                    continue
                mid = msg_id(m)
                if mid in have or mid in seen_now:
                    continue
                seen_now.add(mid)
                snd, to, cc = people(m)
                when = m.SentOn if snd[1] == me else m.ReceivedTime
                ts = dt.datetime(when.year, when.month, when.day, when.hour, when.minute)
                way = "보냄" if snd[1] == me else "받음"
                other = (to[0][0] if to else "") if way == "보냄" else snd[0]
                fname = "%s_%s_%s_%s.msg" % (ts.strftime("%y%m%d_%H%M"), way, safe(other, 12), safe(m.Subject, 50))
                path = os.path.join(out_dir, ts.strftime("%y%m"), fname)
                body, blen = clean_body(m.Body)
                rec = {"id": mid, "at": ts.isoformat(timespec="minutes"), "dir": way, "folder": rel,
                       "from": {"name": snd[0], "addr": snd[1], "role": role(snd[1], c, me)},
                       "to": [{"name": n, "addr": x, "role": role(x, c, me)} for n, x in to],
                       "cc": [{"name": n, "addr": x, "role": role(x, c, me)} for n, x in cc],
                       "subject": m.Subject or "", "topic": getattr(m, "ConversationTopic", "") or "",
                       "attachments": [at.FileName for at in m.Attachments if getattr(at, "Type", 1) == 1],
                       "body": body, "body_len": blen, "msg": os.path.relpath(path, d)}
                new.append((rec, m, path))
                newest = max(newest, ts.isoformat(timespec="minutes"))
        for sf in fo.Folders:                                           # 거래처 폴더 안 하위 폴더도
            walk(sf, rel + "\\" + sf.Name)

    walk(f, f.Name)
    new.sort(key=lambda x: x[0]["at"])
    print("%s - 새 메일 %d통 (since %s, 이미 있는 %d통 제외)" % (os.path.basename(d), len(new), since_dt.date(), len(have)))
    for rec, _, _ in new:
        print("  %s %s %s%s" % (rec["at"].replace("T", " "), rec["dir"], rec["subject"][:60],
                                " [첨부 %d]" % len(rec["attachments"]) if rec["attachments"] else ""))
    if a.dry_run or not new:
        return
    att_n = 0
    with open(idx_p, "a", encoding="utf-8") as fo:
        for k_new, (rec, m, path) in enumerate(new, 1):
            os.makedirs(os.path.dirname(path), exist_ok=True)
            p, k = path, 2
            while os.path.exists(p):
                p = path[:-4] + " (%d).msg" % k
                k += 1
            m.SaveAs(p, OL_MSG_UNICODE)                                 # 사본 저장 - 메일함의 원본은 그대로
            rec["msg"] = os.path.relpath(p, d)
            if rec["dir"] == "받음" and rec["attachments"] and c.get("save_attachments", True):
                dst = os.path.join(attach_dir(d), "%s_#%d_%s_%s" % (os.path.basename(p)[:11], n0 + k_new,
                                                                     safe(rec["from"]["name"], 12), safe(rec["subject"], 40)))
                got = save_attachments(m, dst)
                if got:
                    rec["attach_dir"] = os.path.relpath(dst, d)
                    att_n += len(got)
            fo.write(json.dumps(rec, ensure_ascii=False) + "\n")
    state.update({"last_received": newest, "last_run": dt.datetime.now().isoformat(timespec="seconds"), "count": len(have) + len(new)})
    write_json(st_p, state)
    print("저장 %d통 → %s%s" % (len(new), out_dir, " · 첨부 %d개 → %s" % (att_n, attach_dir(d)) if att_n else ""))


if __name__ == "__main__":
    sys.exit(main())
