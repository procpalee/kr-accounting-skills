# -*- coding: utf-8 -*-
"""builder_history - git 을 쓰지 않는 폴더의 빌더 코드 이력. 빌더가 문서를 저장할 때 자동으로 쌓인다.

저장 위치: <빌더 폴더>/_history/<빌더 이름>/<시각>_<해시8자리>.py, 실행 기록은 <빌더 폴더>/_history/builds.tsv

  python -X utf8 builder_history.py list    <빌더.py>
  python -X utf8 builder_history.py diff    <빌더.py> <판> [<판2>]     판2 를 빼면 지금 파일과 비교
  python -X utf8 builder_history.py restore <빌더.py> <판>             지금 파일을 이력에 남긴 뒤 그 판으로 되돌림
  python -X utf8 builder_history.py which   <문서>                     이 문서를 만든 빌더 판 찾기
판은 list 에 나오는 해시 8자리다.
"""
import datetime
import difflib
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import office_guard as og  # noqa: E402


def _dir(script):
    d, name = os.path.split(os.path.abspath(script))
    return os.path.join(d, "_history", os.path.splitext(name)[0])


def _versions(script):
    d = _dir(script)
    return sorted(f for f in os.listdir(d) if f.endswith(".py")) if os.path.isdir(d) else []


def _find(script, ver):
    hits = [f for f in _versions(script) if f.endswith("_%s.py" % ver) or f == ver]
    if len(hits) != 1:
        raise SystemExit("판을 찾지 못했습니다: %s" % ver)
    return os.path.join(_dir(script), hits[0])


def _read(p):
    with open(p, encoding="utf-8", errors="replace") as f:
        return f.read().splitlines()


def keep(script):
    """지금 파일을 이력에 넣는다(같은 내용이 이미 있으면 건너뜀). 반환: 해시 8자리."""
    sha = og.sha1(script)[:8]
    if not any(f.endswith("_%s.py" % sha) for f in _versions(script)):
        d = _dir(script)
        os.makedirs(d, exist_ok=True)
        shutil.copy2(script, os.path.join(d, "%s_%s.py" % (datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), sha)))
    return sha


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    a = sys.argv[1:]
    if len(a) < 2 or a[0] not in ("list", "diff", "restore", "which"):
        print(__doc__)
        return 2
    cmd, target = a[0], os.path.abspath(a[1])
    if cmd == "list":
        cur = og.sha1(target)[:8] if os.path.isfile(target) else ""
        for f in _versions(target):
            ver = f[:-3].rsplit("_", 1)[1]
            print("%s  %s%s" % (f[:15], ver, "  (지금 파일)" if ver == cur else ""))
        return 0
    if cmd == "diff":
        old = _find(target, a[2])
        new = _find(target, a[3]) if len(a) > 3 else target
        for line in difflib.unified_diff(_read(old), _read(new), os.path.basename(old), os.path.basename(new), lineterm=""):
            print(line)
        return 0
    if cmd == "restore":
        src = _find(target, a[2])
        keep(target)
        shutil.copy2(src, target)
        print("되돌림: %s -> %s" % (os.path.basename(src), os.path.basename(target)))
        return 0
    sha = og.sha1(target)[:8]
    probe, found = os.path.dirname(target), False
    while True:
        log = os.path.join(probe, "_history", "builds.tsv")
        for base in [log] + [os.path.join(probe, x, "_history", "builds.tsv") for x in (os.listdir(probe) if os.path.isdir(probe) else []) if x.startswith("_build")]:
            if os.path.exists(base):
                with open(base, encoding="utf-8") as f:
                    for line in f:
                        p = line.rstrip("\n").split("\t")
                        if len(p) >= 5 and p[4] == sha:
                            print("%s  %s 판 %s  (%s)" % (p[0], p[1], p[2], base))
                            found = True
        up = os.path.dirname(probe)
        if up == probe or found:
            break
        probe = up
    if not found:
        print("이 내용을 만든 빌더 기록을 찾지 못했습니다(사람이 고쳤거나 이력 기능 이전의 파일).")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
