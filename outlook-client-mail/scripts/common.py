# -*- coding: utf-8 -*-
r"""outlook-client-mail 공통 - 설정 읽기 · 거래처 폴더 찾기 · 파일 쓰기.

거래처 루트(설정 root) 구조:
  <root>\_메일분류\                     설정 · 기록(mail_sort.json, Client_Map.tsv, 로그) - VBA 도 여기를 본다
  <root>\[그룹\]<거래처>\               거래처 폴더(이름 앞 'NN. ' 번호는 비교할 때 뗀다)
      메일\YYMM\*.msg                    원본
      첨부\<메일 단위 폴더>\             받은 메일 첨부
      메일 히스토리.md                   Claude 가 쓰는 정리 문서
      _mail\mail.json                    거래처 설정(주소 · 도메인 · 키워드 · Outlook 폴더 연결)
      _mail\mail_index.jsonl             메일 한 통 = 한 줄(번호 = 줄 번호)
      _mail\mail_state.json              마지막 수집 시각 · 히스토리 반영 번호
루트 위치는 환경변수 OCM_ROOT, 없으면 %APPDATA%\outlook-client-mail\config.json 의 root(setup.py 가 만든다).
"""
import datetime as dt
import glob
import json
import os
import re

APP_CFG = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "outlook-client-mail", "config.json")
CONF_DIR = "_메일분류"
META = "_mail"
DEFAULTS = {
    "outlook_root": "거래처",          # Outlook 기본 사서함 최상위 아래 거래처 폴더들의 부모 폴더 이름
    "groups": [],                      # 거래처를 한 단계 묶는 폴더(예: ["01. 고객사", "02. 협력사"]). 비우면 루트 바로 아래가 거래처
    "watch": [],                       # 받은 · 보낸 편지함 말고도 분류할 폴더(서버 규칙으로 먼저 옮겨지는 팀 폴더 등, 사서함 최상위 기준 경로)
    "exclude_subjects": ["그룹에 가입했습니다", "작업이 할당되었습니다", "폴더를 사용자와 공유했습니다", "명함을 보냈습니다"],
    "max_recipients": 15,              # 받는 사람이 이보다 많으면 단체 메일로 보고 주소로 분류하지 않는다
    "team": [],                        # 우리 쪽 사람 주소(내 메일 도메인은 적지 않아도 우리 쪽)
    "since": "",                       # 이 날짜 이전 메일은 모으지 않는다(비우면 setup 한 날)
    "save_dir": "메일",
    "attach_dir": "첨부",
    "history_file": "메일 히스토리.md",
}


def root():
    r = os.environ.get("OCM_ROOT")
    if not r and os.path.exists(APP_CFG):
        r = json.load(open(APP_CFG, encoding="utf-8")).get("root")
    if not r:
        raise SystemExit("거래처 루트가 설정되지 않았습니다 - 먼저 setup.py --root <폴더> 를 실행하세요")
    return os.path.abspath(r)


def conf_dir():
    return os.path.join(root(), CONF_DIR)


def settings():
    p = os.path.join(conf_dir(), "mail_sort.json")
    s = dict(DEFAULTS)
    if os.path.exists(p):
        s.update(json.load(open(p, encoding="utf-8")))
    if not s.get("since"):
        s["since"] = dt.date.today().isoformat()
    return s


def write_text(p, text, encoding="utf-8", newline=None):
    """덮어쓰기 - 숨김 파일이어도(Windows 는 숨김 파일을 'w' 로 열면 권한 오류) 속성은 두고 내용만 바꾼다."""
    if os.path.exists(p):
        with open(p, "r+", encoding=encoding, newline=newline) as f:
            f.seek(0)
            f.write(text)
            f.truncate()
    else:
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding=encoding, newline=newline) as f:
            f.write(text)


def write_json(p, obj, indent=1):
    write_text(p, json.dumps(obj, ensure_ascii=False, indent=indent))


def bare(name):
    """'03. 가나상사' → '가나상사'."""
    return re.sub(r"^\d+\.\s*", "", name).strip()


def _is_client_dir(d):
    b = os.path.basename(d)
    return os.path.isdir(d) and not b.startswith(("_", ".", "~"))


def clients():
    """거래처 폴더 → [{"dir", "name", "group"}] (그룹 순 · 이름 순)."""
    s, r = settings(), root()
    out = []
    groups = s.get("groups") or [""]
    for g in groups:
        base = os.path.join(r, g) if g else r
        if not os.path.isdir(base):
            continue
        for d in sorted(glob.glob(os.path.join(base, "*"))):
            if not _is_client_dir(d):
                continue
            out.append({"dir": d, "name": bare(os.path.basename(d)), "group": g})
    return out


def resolve(arg):
    """거래처 폴더 경로 · 이름(일부) → 폴더 경로. 이름이 정확히 같은 것이 하나면 그것."""
    if os.path.isdir(arg):
        return os.path.abspath(arg)
    cs = clients()
    exact = [c for c in cs if c["name"] == arg]
    hit = exact if len(exact) == 1 else [c for c in cs if arg in c["name"]]
    if len(hit) != 1:
        raise SystemExit("거래처를 하나로 못 찾음(%s): %s" % (arg, [c["name"] for c in hit] or "없음"))
    return hit[0]["dir"]


def meta(d, *parts):
    return os.path.join(d, META, *parts)


def mail_cfg_path(d):
    return meta(d, "mail.json")


def load_mail_cfg(d):
    p = mail_cfg_path(d)
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else None


def new_mail_cfg(name):
    s = settings()
    return {"since": s["since"], "company": {"addresses": [], "domains": []}, "others": {"addresses": [], "domains": []},
            "keywords": [name]}


def hide(p):
    """_mail 폴더를 숨김으로(윈도우만, 실패해도 그만)."""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetFileAttributesW(p, 0x02)
        except Exception:  # noqa: BLE001
            pass


def save_dir(d):
    return os.path.join(d, settings()["save_dir"])


def attach_dir(d):
    return os.path.join(d, settings()["attach_dir"])


def history_path(d):
    return os.path.join(d, settings()["history_file"])
