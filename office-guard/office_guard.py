# -*- coding: utf-8 -*-
"""office_guard - 코드가 사람이 고친 Office 문서(xlsx·xlsm·docx·pptx·dsd)를 덮어쓰지 못하게 하는 공용 판단·저장 함수.

표준 라이브러리만 쓴다(훅의 파이썬에는 openpyxl 등이 없다).

판단 순서(judge):
  파일 없음 -> 허용 / 임시 폴더·_backup·_xlsync·_docsync 안 -> 허용
  현재 내용이 코드가 저장한 내용(장부)과 같음 -> 허용
  현재 내용이 기준본(_xlsync·_docsync base.json)과 같음 -> 허용
  git 에 커밋된 그대로 -> 허용
  같은 파일을 불러와 일부만 고친 저장(부분 편집) -> 백업 뒤 허용, 코드 소유로 치지 않음
  그 밖 -> 거부

명령:
  python -X utf8 office_guard.py status <파일>
  python -X utf8 office_guard.py check  <파일>     거부 대상이면 종료코드 1 (COM·외부 도구 저장 전에 호출)
  python -X utf8 office_guard.py accept <파일>     사용자가 덮어써도 된다고 확인한 뒤에만
  python -X utf8 office_guard.py revoke <파일>     장부 기록을 내려 다시 사람 파일로 본다(승인 취소, 잘못 올라간 기록 정리)
  python -X utf8 office_guard.py handover <파일|폴더> [--note 메모]   사람 파일을 정본으로(사용자가 말했을 때만)
  python -X utf8 office_guard.py handover <파일|폴더> --back          코드를 다시 정본으로(사용자 확인 뒤에만)
"""
import datetime
import hashlib
import json
import os
import re
import shutil
import sys

GUARDED = (".xlsx", ".xlsm", ".docx", ".pptx", ".dsd")
EXEMPT_DIRS = {"_backup", "_xlsync", "_docsync", "_history", "__pycache__"}
TOOL = "~/.claude/skills/office-guard/office_guard.py"


class OverwriteRefused(PermissionError):
    """사람이 고친 파일에 새로 만든 내용을 저장하려 할 때."""


def state_dir():
    return os.environ.get("OFFICE_GUARD_STATE") or os.path.join(os.path.expanduser("~"), ".claude", "office-guard")


def is_guarded(path):
    name = os.path.basename(str(path)).lower()
    return name.endswith(GUARDED) and not name.startswith("~$")


def sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def stamp(path):
    st = os.stat(path)
    return (st.st_mtime_ns, st.st_size)


def same_path(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


# ── 장부: 코드가 저장한 내용의 해시. 경로를 쓰지 않으므로 폴더를 옮겨도 그대로 맞고 고객 이름도 남지 않는다.
def _ledger_file(sha):
    return os.path.join(state_dir(), "ledger", sha[:2], sha)


def is_own(sha):
    return os.path.exists(_ledger_file(sha))


def remember(sha):
    p = _ledger_file(sha)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write(datetime.datetime.now().isoformat(timespec="seconds"))


def forget(sha):
    try:
        os.remove(_ledger_file(sha))
    except OSError:
        pass


# 사람 파일을 코드가 다른 경로에 복사했을 때는 그 경로에서만 코드 소유로 친다 - 원본은 계속 사람 파일이다.
def _at_key(sha, path):
    where = hashlib.sha1(os.path.normcase(os.path.abspath(path)).encode("utf-8")).hexdigest()[:16]
    return "%s@%s" % (sha, where)


def is_own_at(sha, path):
    return os.path.exists(_ledger_file(_at_key(sha, path)))


def remember_at(sha, path):
    remember(_at_key(sha, path))


def forget_at(sha, path):
    forget(_at_key(sha, path))


def _temp_roots():
    import tempfile
    roots = {tempfile.gettempdir()}
    for k in ("TEMP", "TMP"):
        if os.environ.get(k):
            roots.add(os.environ[k])
    return [os.path.normcase(os.path.abspath(r)) for r in roots]


def exempt_place(path):
    parts = {x.lower() for x in os.path.normpath(path).split(os.sep)}
    if parts & EXEMPT_DIRS:
        return True
    if os.environ.get("OFFICE_GUARD_INCLUDE_TEMP"):
        return False
    p = os.path.normcase(os.path.abspath(path))
    return any(p.startswith(r + os.sep) for r in _temp_roots())


def baseline_file(path):
    """이 파일의 기준본 base.json 경로(없어도 경로는 돌려준다)."""
    d, name = os.path.split(os.path.abspath(path))
    if name.lower().endswith((".xlsx", ".xlsm")):
        return os.path.join(d, "_xlsync", os.path.splitext(name)[0], "base.json")
    return os.path.join(d, "_docsync", name, "base.json")


def baseline_sha(path):
    b = baseline_file(path)
    if not os.path.exists(b):
        return None
    with open(b, encoding="utf-8", errors="ignore") as f:
        head = f.read(65536)
    m = re.search(r'"sha1":\s*"([0-9a-f]{40})"', head)
    if m:
        return m.group(1)
    with open(b, encoding="utf-8") as f:
        return json.load(f).get("meta", {}).get("sha1")


def git_clean(path):
    """git 에 커밋된 그대로면 True - 덮어써도 git 에서 되살릴 수 있다."""
    d = os.path.dirname(os.path.abspath(path))
    probe = d
    while not os.path.exists(os.path.join(probe, ".git")):
        up = os.path.dirname(probe)
        if up == probe:
            return False
        probe = up
    import subprocess
    name = os.path.basename(path)
    try:
        run = lambda *a: subprocess.run(["git", "-C", d, *a], capture_output=True, timeout=5).returncode
        return run("ls-files", "--error-unmatch", "--", name) == 0 and run("diff", "--quiet", "HEAD", "--", name) == 0
    except Exception:
        return False


def source_of(path):
    """불러온 문서에 달아 두는 표식: (경로, 불러올 때의 수정시각·크기)."""
    p = os.path.abspath(path)
    return (p, stamp(p))


# ── 정본 전환: 사람 파일이 정본이 된 시점. 건 폴더에 한 파일에 한 줄씩 남는다(마지막 줄이 현재 상태).
HANDOVER_FILE = "정본전환.tsv"


def _handover_home(path):
    """기록 위치: 위로 올라가며 처음 만나는 _engagement, 없으면 파일 옆 _xlsync. 반환: (기록 파일, 기준 폴더)."""
    d = os.path.dirname(os.path.abspath(path))
    probe = d
    while True:
        if os.path.isdir(os.path.join(probe, "_engagement")):
            return os.path.join(probe, "_engagement", HANDOVER_FILE), probe
        up = os.path.dirname(probe)
        if up == probe:
            return os.path.join(d, "_xlsync", HANDOVER_FILE), d
        probe = up


def _handover_key(path, base):
    return os.path.relpath(os.path.abspath(path), base).replace("\\", "/")


def handover_state(path):
    """이 파일의 마지막 전환 기록. 반환: (상태, 시각) 또는 None. 상태는 '사람' 또는 '코드'."""
    rec, base = _handover_home(path)
    if not os.path.exists(rec):
        return None
    key, last = _handover_key(path, base), None
    with open(rec, encoding="utf-8") as f:
        for line in f:
            parts = line.rstrip("\n").split("\t")
            if len(parts) >= 3 and parts[1] == key:
                last = (parts[2], parts[0])
    return last


def _handover_targets(path):
    if os.path.isdir(path):
        out = []
        for d, subs, names in os.walk(path):
            subs[:] = [s for s in subs if s.lower() not in EXEMPT_DIRS]
            out += [os.path.join(d, n) for n in sorted(names) if is_guarded(n)]
        return out
    return [path]


def set_handover(path, state, note=""):
    """파일(또는 폴더 안의 기존 문서)의 정본 상태를 기록한다. 반환: 기록한 파일 수."""
    targets = _handover_targets(os.path.abspath(path))
    now = datetime.datetime.now().isoformat(timespec="seconds")
    for t in targets:
        rec, base = _handover_home(t)
        os.makedirs(os.path.dirname(rec), exist_ok=True)
        with open(rec, "a", encoding="utf-8") as f:
            f.write("%s\t%s\t%s\t%s\n" % (now, _handover_key(t, base), state, note.replace("\t", " ").replace("\n", " ")))
    return len(targets)


def is_handed_over(path):
    st = handover_state(path)
    return bool(st) and st[0] == "사람"


def judge(path, source=None):
    """저장해도 되는지 판단. 반환: dict(ok, why, own, backup). own=저장 뒤 장부에 올릴지, backup=저장 전 사본을 남길지."""
    p = os.path.abspath(path)
    if not is_guarded(p) or not os.path.isfile(p):
        return dict(ok=True, why="new", own=True, backup=False)
    if exempt_place(p):
        return dict(ok=True, why="place", own=True, backup=False)
    if is_handed_over(p):                       # 전환 뒤에는 코드가 만든 그대로여도 다시 만들지 못한다
        if source and same_path(source[0], p):
            if tuple(source[1]) == stamp(p):
                return dict(ok=True, why="edit", own=False, backup=True)
            return dict(ok=False, why="changed_since_load", own=False, backup=False)
        return dict(ok=False, why="handover", own=False, backup=False)
    cur = sha1(p)
    if is_own(cur) or is_own_at(cur, p):
        return dict(ok=True, why="code", own=True, backup=False)
    base = baseline_sha(p)
    if base == cur:
        return dict(ok=True, why="baseline", own=True, backup=False)
    if git_clean(p):
        return dict(ok=True, why="git", own=True, backup=False)
    if source and same_path(source[0], p):
        if tuple(source[1]) == stamp(p):
            return dict(ok=True, why="edit", own=False, backup=True)
        return dict(ok=False, why="changed_since_load", own=False, backup=False)
    return dict(ok=False, why="changed" if base else "unknown", own=False, backup=False)


def owned_content(path):
    """이 파일의 내용이 코드가 만든 그대로인지(장부·기준본·git). 위치는 보지 않는다 - 복사가 내용의 주인을 바꾸지 않게 하는 데 쓴다."""
    p = os.path.abspath(path)
    if not os.path.isfile(p):
        return False
    cur = sha1(p)
    if is_own(cur):
        return True
    if not is_guarded(p):
        return False
    return baseline_sha(p) == cur or git_clean(p)


REASONS = {
    "changed": "기준본을 남긴 뒤 내용이 바뀌었습니다(사람이 고친 것으로 봅니다).",
    "unknown": "코드가 만든 기록도 기준본도 없는 기존 파일입니다(사람이 고쳤는지 알 수 없어 고친 것으로 봅니다).",
    "changed_since_load": "코드가 불러온 뒤 파일이 다시 저장됐습니다(그사이 사람이 고쳤을 수 있습니다).",
}


def refusal(path, v):
    p = os.path.abspath(path)
    if v["why"] == "handover":
        return (
            "[저장 거부] 정본이 사람 파일로 넘어간 문서입니다: %s\n"
            "  전환 시각: %s\n"
            "  할 일:\n"
            "   1. 이 파일을 빌더로 다시 만들지 않는다. 결과를 비교하려면 임시 폴더로 출력한다.\n"
            "   2. 몇 칸만 고칠 일이면 그 파일을 불러와 고친 뒤 같은 경로에 저장한다(부분 편집은 허용).\n"
            "   3. 사용자가 코드를 다시 정본으로 하겠다고 말했을 때만: python -X utf8 %s handover \"%s\" --back"
            % (p, (handover_state(p) or ("", ""))[1], TOOL, p))
    return (
        "[저장 거부] 덮어쓰지 않았습니다: %s\n"
        "  이유: %s\n"
        "  할 일:\n"
        "   1. 출력 경로를 임시 폴더나 다른 이름으로 바꿔 저장한다.\n"
        "   2. 사람이 고친 내용을 코드에 옮기려면 excel-sync(diff·verify)로 비교한다.\n"
        "   3. 일부만 고칠 일이면 그 파일을 불러와 고친 뒤 같은 경로에 저장한다(부분 편집은 허용).\n"
        "   4. 사용자가 덮어써도 된다고 확인했을 때만: python -X utf8 %s accept \"%s\""
        % (p, REASONS.get(v["why"], v["why"]), TOOL, p))


def backup(path):
    """<폴더>/_backup/<이름>_bak_시각.확장자 로 사본. 직전 사본과 내용이 같으면 건너뛴다. 반환: 사본 경로 또는 None."""
    d = os.path.join(os.path.dirname(path), "_backup")
    os.makedirs(d, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    prev = sorted(f for f in os.listdir(d) if f.startswith(base + "_bak_") and f.endswith(ext))
    if prev and sha1(os.path.join(d, prev[-1])) == sha1(path):
        return None
    tag = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = os.path.join(d, "%s_bak_%s%s" % (base, tag, ext))
    n = 1
    while os.path.exists(dst):
        n += 1
        dst = os.path.join(d, "%s_bak_%s_%d%s" % (base, tag, n, ext))
    with open(path, "rb") as a, open(dst, "xb") as b:   # 가드가 감싼 복사를 쓰지 않는다 - 사본의 내용은 장부에 올리지 않는다
        shutil.copyfileobj(a, b)
    try:
        shutil.copystat(path, dst)
    except OSError:
        pass
    return dst


def before_write(path, source=None):
    """저장 직전 호출. 거부면 OverwriteRefused, 허용이면 판단 결과를 돌려준다."""
    v = judge(path, source)
    if not v["ok"]:
        raise OverwriteRefused(refusal(path, v))
    if v["backup"]:
        backup(os.path.abspath(path))
    return v


def after_write(path, v):
    """저장 직후 호출. 코드 소유 내용이면 장부에 올리고, 사람 내용을 고친 저장이면 내린다."""
    p = os.path.abspath(path)
    if not is_guarded(p) or not os.path.isfile(p):
        return None
    cur = sha1(p)
    if v["own"] == "at":                        # 사람 파일의 사본: 이 경로에서만 코드 소유
        remember_at(cur, p)
    elif v["own"]:
        remember(cur)
    else:
        forget(cur)
        forget_at(cur, p)
    return cur


def _source_of_object(obj):
    src = getattr(obj, "_og_src", None)
    if src:
        return src
    for attr in ("package", "part"):                # python-docx·python-pptx 는 패키지에 표식이 있다
        try:
            inner = getattr(obj, attr)
            src = getattr(inner, "_og_src", None) or getattr(getattr(inner, "package", None), "_og_src", None)
        except Exception:
            src = None
        if src:
            return src
    return None


def safe_save(obj, path, on_conflict="refuse", source=None):
    """obj.save(path) 를 판단과 함께 실행. 반환: 실제로 저장한 경로.

    on_conflict="refuse"  : 사람이 고친 파일이면 OverwriteRefused
    on_conflict="sidecar" : 사람이 고친 파일이면 옆에 <이름>_재생성.확장자 로 저장
    source                : 같은 파일을 불러와 고칠 때 불러오기 전에 얻은 source_of(경로). 런타임 가드가 없는 환경의 부분 편집용
    """
    path = os.path.abspath(path)
    source = source or _source_of_object(obj)
    try:
        v = before_write(path, source)
    except OverwriteRefused:
        if on_conflict != "sidecar":
            raise
        base, ext = os.path.splitext(path)
        path = base + "_재생성" + ext
        if os.path.exists(path) and not judge(path)["ok"]:
            path = "%s_재생성_%s%s" % (base, datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), ext)
        v = before_write(path)
    obj.save(path)
    after_write(path, v)
    return path


def _in_git(path):
    probe = os.path.dirname(os.path.abspath(path))
    while True:
        if os.path.exists(os.path.join(probe, ".git")):
            return True
        up = os.path.dirname(probe)
        if up == probe:
            return False
        probe = up


def record_build(script, outputs, modules=()):
    """문서를 저장한 빌더의 코드를 빌더 옆 _history 에 남기고 실행 기록 한 줄을 쓴다.

    git 저장소·임시 폴더·설치된 패키지의 스크립트는 남기지 않는다. 내용이 같은 판은 다시 복사하지 않는다.
    """
    if not script or not str(script).lower().endswith(".py") or not os.path.isfile(script) or not outputs:
        return
    script = os.path.abspath(script)
    if "site-packages" in script or exempt_place(script) or _in_git(script):
        return
    d = os.path.dirname(script)
    now = datetime.datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    versions = []
    for p in dict.fromkeys([script] + [os.path.abspath(m) for m in modules if m and os.path.isfile(m)]):
        sha = sha1(p)[:8]
        stem = os.path.splitext(os.path.basename(p))[0]
        h = os.path.join(d, "_history", stem)
        os.makedirs(h, exist_ok=True)
        if not any(f.endswith("_%s.py" % sha) for f in os.listdir(h)):
            shutil.copy2(p, os.path.join(h, "%s_%s.py" % (ts, sha)))
        versions.append((os.path.basename(p), sha))
    log = os.path.join(d, "_history", "builds.tsv")
    with open(log, "a", encoding="utf-8") as f:
        for out in outputs:
            if os.path.isfile(out):
                f.write("%s\t%s\t%s\t%s\t%s\t%s\n" % (now.isoformat(timespec="seconds"), versions[0][0], versions[0][1], os.path.basename(out), sha1(out)[:8],
                                                  ",".join("%s:%s" % v for v in versions[1:])))


WHY_TEXT = {
    "new": "파일이 없어 새로 저장합니다.",
    "place": "임시·기록 폴더 안이라 검사하지 않습니다.",
    "code": "코드가 저장한 그대로입니다. 다시 만들어도 됩니다.",
    "baseline": "기준본과 같습니다. 다시 만들어도 됩니다.",
    "git": "git 에 커밋된 그대로입니다. 다시 만들어도 됩니다.",
}


def main(argv=None):
    sys.stdout.reconfigure(encoding="utf-8")
    argv = sys.argv[1:] if argv is None else argv
    if len(argv) >= 2 and argv[0] == "handover":
        import argparse
        ap = argparse.ArgumentParser(prog="office_guard.py handover")
        ap.add_argument("path")
        ap.add_argument("--back", action="store_true")
        ap.add_argument("--note", default="")
        a = ap.parse_args(argv[1:])
        n = set_handover(a.path, "코드" if a.back else "사람", a.note)
        print("%s: 문서 %d개 - %s" % ("코드를 다시 정본으로" if a.back else "사람 파일을 정본으로", n, os.path.abspath(a.path)))
        return 0 if n else 2
    if len(argv) != 2 or argv[0] not in ("status", "check", "accept", "revoke"):
        print(__doc__)
        return 2
    cmd, path = argv[0], os.path.abspath(argv[1])
    if cmd == "revoke":
        if not os.path.isfile(path):
            print("파일 없음: " + path)
            return 2
        cur = sha1(path)
        had = is_own(cur) or is_own_at(cur, path)
        forget(cur)
        forget_at(cur, path)
        v = judge(path)
        print("%s - %s" % ("장부 기록을 내렸습니다" if had else "장부에 기록이 없습니다", path))
        print("  지금 판정: " + ("다시 만들 수 있음(%s)" % v["why"] if v["ok"] else "사람 파일로 보고 거부"))
        return 0
    if cmd == "accept":
        if not os.path.isfile(path):
            print("파일 없음: " + path)
            return 2
        kept = None
        if is_guarded(path) and not exempt_place(path) and not owned_content(path):
            kept = backup(path)                 # 승인 뒤 빌더가 덮어쓰기 전에 사람이 고친 판을 남긴다
        remember(sha1(path))
        print("승인 기록: 지금 내용은 코드가 다시 만들어도 되는 상태로 칩니다 - " + path)
        if kept:
            print("  승인 전 내용의 사본: " + kept)
        return 0
    v = judge(path)
    if v["ok"]:
        print(WHY_TEXT.get(v["why"], v["why"]) + " - " + path)
        return 0
    print(refusal(path, v))
    return 1


if __name__ == "__main__":
    sys.exit(main())
