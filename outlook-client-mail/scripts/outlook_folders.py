# -*- coding: utf-8 -*-
r"""거래처 폴더 ↔ Outlook 거래처 폴더를 연결하고 계속 맞춘다.

  python -X utf8 outlook_folders.py            미리보기(아무것도 바꾸지 않음)
  python -X utf8 outlook_folders.py --apply    맞추기(mail_sync_all.py 가 매번 실행 - 바뀐 것만 출력)

연결 기록 = 거래처 폴더 안 _mail\mail.json 의 "outlook": {"entry_id", "store_id", "path", "active"}
  - EntryID 는 Outlook 폴더 이름을 바꾸거나 옮겨도 그대로다 → Outlook 쪽에서 이름 · 위치를 바꿔도 끊기지 않는다.
  - mail.json 은 거래처 폴더 안에 있다 → 로컬 폴더 이름 · 번호를 바꿔도 연결이 따라간다.
  - 이름 비교는 연결이 없을 때 처음 한 번만 쓴다.
처리:
  연결 없는 거래처 → outlook_root 아래 연결 안 된 폴더 중 이름이 같거나 '이름_' · '이름 (' · '이름(' 로 시작하는 폴더가
                      1개면 연결, 0개면 <outlook_root>\[그룹\]<폴더 이름> 을 만들어 연결, 2개 이상이면 보고만
  연결된 Outlook 폴더를 outlook_root 밖(예: 보관)으로 옮김 → active=false(자동 분류 · 저장 안 함)
  연결된 Outlook 폴더가 없어짐(지워짐 · 다른 사서함 · PST) → 보고만
  어느 거래처와도 연결 안 된 Outlook 폴더 → 보고만(건드리지 않음)
메일 설정(mail.json)이 없는 거래처는 빈 기본값으로 만든다(주소는 VBA [거래처로 분류] '기억' 이나 직접 추가).
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import bare, clients, hide, load_mail_cfg, mail_cfg_path, meta, new_mail_cfg, settings, write_json  # noqa: E402


def retry(fn, tries=12, wait=10):
    """Outlook 이 막 켜져 '오프라인' 이면 잠깐 기다렸다 다시."""
    import time
    for i in range(tries):
        try:
            return fn()
        except Exception as ex:  # noqa: BLE001
            if i == tries - 1 or ("오프라인" not in str(ex) and "offline" not in str(ex).lower()):
                raise
            time.sleep(wait)


def rel_of(f, root_path):
    fp = f.FolderPath
    i = fp.lower().find(root_path.lower() + "\\")
    return fp[i + len(root_path) + 1:] if i >= 0 else fp


def leaves(croot, groups):
    """outlook_root 아래 거래처 폴더(그룹 폴더는 안으로 들어감) → [폴더]."""
    gs = {g.lower() for g in groups}
    out = []

    def walk(f):
        for sf in f.Folders:
            if sf.Name.lower() in gs:
                walk(sf)
            else:
                out.append(sf)
    walk(croot)
    return out


def run(apply):
    import win32com.client
    s = settings()
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    store = ns.DefaultStore
    top = retry(lambda: store.GetRootFolder())
    msgs = []
    try:
        croot = top.Folders.Item(s["outlook_root"])
    except Exception:  # noqa: BLE001
        if not apply:
            return ["만들기 - Outlook 최상위 폴더 '%s'(--apply 때 만듦)" % s["outlook_root"]]
        croot = retry(lambda: top.Folders.Add(s["outlook_root"]))
        msgs.append("만들기 - Outlook 최상위 폴더 '%s'" % s["outlook_root"])
    root_path = croot.FolderPath.split("\\", 3)[-1]
    cs = clients()
    for c in cs:
        c["cfg"] = load_mail_cfg(c["dir"])
    claimed = {c["cfg"]["outlook"]["entry_id"] for c in cs if c["cfg"] and c["cfg"].get("outlook", {}).get("entry_id")}
    lv = leaves(croot, s.get("groups") or [])

    def by_id(eid):
        try:
            return ns.GetFolderFromID(eid, store.StoreID)
        except Exception:  # noqa: BLE001
            return None

    def parent_for(group):
        f = croot
        if group:
            try:
                f = f.Folders.Item(group)
            except Exception:  # noqa: BLE001
                f = retry(lambda: croot.Folders.Add(group))
        return f

    for c in cs:
        cfg, leaf = c["cfg"], os.path.basename(c["dir"])
        target = (c["group"] + "\\" if c["group"] else "") + leaf
        link = (cfg or {}).get("outlook", {}).get("entry_id")
        f = None
        if link:
            f = by_id(link)
            if f is None:
                msgs.append("연결 끊김(Outlook 폴더 없음) - %s: 지워졌거나 다른 사서함 · PST 로 옮겨짐. "
                            "폴더를 되살리거나 mail.json 의 outlook 을 지우면 다음에 다시 연결" % target)
                continue
        else:
            hits = [x for x in lv if x.EntryID not in claimed and
                    (bare(x.Name) == c["name"] or re.match(re.escape(c["name"]) + r"(_| ?\()", bare(x.Name)))]
            if len(hits) > 1:
                msgs.append("확인 필요 - '%s' 이름 폴더 %d개: %s" % (c["name"], len(hits), ", ".join(rel_of(x, root_path) for x in hits)))
                continue
            if hits:
                f = hits[0]
                msgs.append("연결 - %s ↔ %s" % (leaf, rel_of(f, root_path)))
            else:
                msgs.append("만들기 - %s" % target)
                if apply:
                    f = retry(lambda: parent_for(c["group"]).Folders.Add(leaf))
        if f is not None and apply:
            if cfg is None:
                cfg = new_mail_cfg(c["name"])
                msgs.append("메일 설정 만듦 - %s(거래처 주소 비어 있음)" % leaf)
            active = f.FolderPath.lower().startswith(croot.FolderPath.lower() + "\\")
            new = {"entry_id": f.EntryID, "store_id": store.StoreID, "path": rel_of(f, root_path), "active": active}
            if cfg.get("outlook") != new:
                if cfg.get("outlook", {}).get("active", True) and not active:
                    msgs.append("보관 - %s: Outlook 폴더가 '%s' 밖으로 옮겨져 자동 분류 · 저장에서 뺌" % (leaf, s["outlook_root"]))
                cfg["outlook"] = new
                os.makedirs(meta(c["dir"]), exist_ok=True)
                hide(meta(c["dir"]))
                write_json(mail_cfg_path(c["dir"]), cfg, 2)
        if f is not None:
            claimed.add(f.EntryID)
    for x in leaves(croot, s.get("groups") or []):
        if x.EntryID not in claimed:
            msgs.append("연결 안 된 Outlook 폴더 - %s(어느 거래처와도 연결 안 됨, 그대로 둠)" % rel_of(x, root_path))
    return msgs


def main():
    apply = "--apply" in sys.argv
    msgs = run(apply)
    print("\n".join(msgs) if msgs else "거래처 폴더와 Outlook 폴더가 모두 연결 · 일치합니다")
    if msgs and not apply:
        print("(미리보기 - 바꾸려면 --apply)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
