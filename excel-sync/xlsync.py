# -*- coding: utf-8 -*-
"""xlsync - 엑셀을 사람이 직접 고치고 저장하면, 기준본과 결정론적으로 비교해 바뀐 점과 고칠 코드 위치를 찾는다.

  python -X utf8 xlsync.py snap   <엑셀> [--prov 기록.json] [--label 설명]
        기준본 저장: <엑셀 폴더>/_xlsync/<이름>/base.json(+원본 사본, 생성 기록). 이전 기준본은 history/ 로.
        새 기준본을 임시 파일로 다 만든 뒤에만 바꾼다 - 도중에 실패해도 기존 기준본은 그대로.
  python -X utf8 xlsync.py status <엑셀>
        기준본 이후 파일이 바뀌었는지(해시)만 확인.
  python -X utf8 xlsync.py diff   <엑셀> [--base 기준(.json/.xlsx)] [--code 코드폴더 ...] [--ignore style,width,...]
        기준본 -> 현재 차이를 표로 출력하고 _xlsync/<이름>/diff_<시각>.md·.json 에 저장. 생성 기록이 있으면 칸마다 '쓴 코드 줄'을 붙인다.
  python -X utf8 xlsync.py verify <기대 엑셀(사람이 고친 것)> <실제 엑셀(코드로 다시 만든 것)> [--ignore ...]
        차이 0 이면 종료코드 0, 남은 차이가 있으면 표를 출력하고 1.
  python -X utf8 xlsync.py record --prov 기록.json -- <생성 스크립트.py> [인자 ...]
        스크립트를 실행하며 openpyxl 이 칸·열·병합·유효성·이름을 쓴 코드 줄을 기록.
  python -X utf8 xlsync.py overrides <diff_...json> --out overrides.json [--kinds value,style,width,height,merge]
        칸 단위 수정을 덮어쓰기 목록으로 저장(생성 코드가 저장 직전에 xlsync_overrides.apply_overrides 로 적용).

엑셀이 저장하지 않은 파일(openpyxl 산출물)은 비교 전에 엑셀로 한 번 저장한 사본으로 읽는다(원본은 그대로). --no-normalize 로 끔.
"""
import argparse
import datetime
import hashlib
import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import xlsync_overrides as ov  # noqa: E402
import xlsync_trace as trace  # noqa: E402
from xlsync_compare import KINDS, compare  # noqa: E402
from xlsync_excel import normalized  # noqa: E402
from xlsync_report import attach, code_index, markdown  # noqa: E402
from xlsync_snapshot import load_or_snap, snapshot  # noqa: E402


def store_dir(path):
    d = os.path.join(os.path.dirname(path), "_xlsync", os.path.splitext(os.path.basename(path))[0])
    os.makedirs(d, exist_ok=True)
    return d


def sha1(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def stamp():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def read_snap(path, normalize=True, label=""):
    if path.lower().endswith(".json"):
        return load_or_snap(path)
    src, resaved = normalized(path, workdir=tempfile.mkdtemp(prefix="xlsync_")) if normalize else (path, False)
    s = snapshot(src, label=label)
    s["meta"].update(file=os.path.abspath(path), sha1=sha1(path), normalized=resaved,
                     mtime=datetime.datetime.fromtimestamp(os.stat(path).st_mtime).isoformat(timespec="seconds"))
    return s


def cmd_snap(a):
    path = os.path.abspath(a.file)
    d = store_dir(path)
    base = os.path.join(d, "base.json")
    copy = "base" + os.path.splitext(path)[1]
    # 새 기준본(스냅샷·원본 사본·생성 기록)을 임시 파일(_tmp_*)로 끝까지 만든 뒤에만 옛 기준본을 history/ 로 보낸다.
    # 읽기·쓰기 도중 실패하면 임시 파일만 지우고 멈춘다 - 기존 base.json·사본·prov.json 은 그대로.
    tmp = {n: os.path.join(d, "_tmp_" + n) for n in ("base.json", copy, "prov.json")}
    try:
        s = read_snap(path, not a.no_normalize, a.label)
        entry = trace.load(a.prov, path) if a.prov else None
        with open(tmp["base.json"], "w", encoding="utf-8") as f:
            json.dump(s, f, ensure_ascii=False)
        shutil.copy2(path, tmp[copy])
        if entry is not None:
            shutil.copy2(a.prov, tmp["prov.json"])
    except BaseException:
        for p in tmp.values():
            if os.path.exists(p):
                os.remove(p)
        raise
    if os.path.exists(base):
        h = os.path.join(d, "history")
        os.makedirs(h, exist_ok=True)
        t = stamp()
        shutil.copy2(base, os.path.join(h, "base_%s.json" % t))       # 복사한 뒤 아래에서 교체 - base.json 이 비는 순간이 없다
        if os.path.exists(os.path.join(d, "prov.json")):
            shutil.move(os.path.join(d, "prov.json"), os.path.join(h, "prov_%s.json" % t))
    for old in os.listdir(d):
        if old.startswith("base.") and not old.endswith(".json") and old != copy:
            os.remove(os.path.join(d, old))
    for n in dict.fromkeys(["base.json", copy] + (["prov.json"] if entry is not None else [])):      # .json 을 snap 하면 copy 도 base.json
        os.replace(tmp[n], os.path.join(d, n))
    if a.prov and entry is None:
        print("[경고] 기록 파일에 저장한 통합문서가 없음: " + a.prov)
    sites = len(entry["sites"]) if entry is not None else 0
    cells = sum(len(x.get("cells", {})) for x in s["sheets"])
    print("기준본 저장: %s" % base)
    print("  시트 %d · 칸 %d · 서식 %d · 체크박스 %d · VBA 모듈 %d · 이름 %d%s%s" % (
        len(s["sheets"]), cells, len(s["styles"]), sum(len(x.get("checkbox", {})) for x in s["sheets"]),
        len(s.get("vba") or {}), len(s["names"]), (" · 코드 위치 %d" % sites) if sites else "",
        " · 엑셀 재저장 사본으로 읽음" if s["meta"].get("normalized") else ""))


def cmd_status(a):
    path = os.path.abspath(a.file)
    base = os.path.join(store_dir(path), "base.json")
    if not os.path.exists(base):
        print("기준본 없음 - 먼저 snap")
        return 2
    meta = json.load(open(base, encoding="utf-8"))["meta"]
    same = meta.get("sha1") == sha1(path)
    print("기준본 %s (%s) · 현재 파일 %s" % (meta.get("made"), meta.get("label") or "", "같음" if same else "바뀜 - diff 필요"))
    return 0 if same else 1


def _ignore(s):
    ig = [x.strip() for x in (s or "").split(",") if x.strip()]
    bad = [x for x in ig if x not in KINDS]
    if bad:
        raise SystemExit("--ignore 에 모르는 종류: %s (가능: %s)" % (bad, ", ".join(KINDS)))
    return ig


def cmd_diff(a):
    path = os.path.abspath(a.file)
    d = store_dir(path)
    base = a.base or os.path.join(d, "base.json")
    if not os.path.exists(base):
        print("기준본 없음: %s - 먼저 'snap' 으로 기준본을 저장" % base)
        return 2
    A = read_snap(base, not a.no_normalize)
    B = read_snap(path, not a.no_normalize)
    res = compare(A, B, _ignore(a.ignore))
    prov = trace.load(os.path.join(d, "prov.json"), path) if (not a.base and os.path.exists(os.path.join(d, "prov.json"))) else None
    roots = a.code or ([prov["root"]] if prov and prov.get("root") else [])
    attach(res, prov, code_index(roots), base=roots[0] if roots else os.path.dirname(path))
    md = markdown(res)
    t = stamp()
    cur = os.path.join(d, "cur_%s.json" % t)
    with open(cur, "w", encoding="utf-8") as f:
        json.dump(B, f, ensure_ascii=False)
    with open(os.path.join(d, "diff_%s.json" % t), "w", encoding="utf-8") as f:
        json.dump({"result": res, "cur_snapshot": cur, "prov": bool(prov), "code_roots": roots}, f, ensure_ascii=False, indent=1)
    with open(os.path.join(d, "diff_%s.md" % t), "w", encoding="utf-8") as f:
        f.write(md)
    print(md)
    print("저장: %s (.md·.json) · 총 %d건" % (os.path.join(d, "diff_%s" % t), res["summary"]["total"]))
    return 0


def cmd_verify(a):
    A = read_snap(os.path.abspath(a.expected), not a.no_normalize)
    B = read_snap(os.path.abspath(a.actual), not a.no_normalize)
    res = compare(A, B, _ignore(a.ignore))
    if res["summary"]["total"] == 0:
        print("일치: 코드로 다시 만든 파일이 기대 파일과 같음(무시: %s)" % (a.ignore or "없음"))
        for n in res.get("notes", []):
            print("  참고: " + n)
        return 0
    print(markdown(res, title="남은 차이(기대 -> 실제)"))
    return 1


def cmd_overrides(a):
    d = json.load(open(a.diff, encoding="utf-8"))
    cur = json.load(open(d["cur_snapshot"], encoding="utf-8"))
    kinds = tuple(x.strip() for x in a.kinds.split(",") if x.strip())
    entries, skipped = ov.make(d["result"], cur, kinds)
    n = ov.save(entries, a.out, note="xlsync overrides - %s" % os.path.basename(a.diff))
    print("덮어쓰기 목록: %s · 이번 항목 %d · 전체 %d" % (a.out, len(entries), n))
    for s in skipped:
        print("  [직접 반영] " + s)
    return 0


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    argv = sys.argv[1:]
    rest = []
    if "--" in argv:
        i = argv.index("--")
        argv, rest = argv[:i], argv[i + 1:]
    ap = argparse.ArgumentParser(description="엑셀 수정 추적·코드 반영 도구")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("snap")
    p.add_argument("file")
    p.add_argument("--prov")
    p.add_argument("--label", default="")
    p.add_argument("--no-normalize", action="store_true")
    p = sub.add_parser("status")
    p.add_argument("file")
    p = sub.add_parser("diff")
    p.add_argument("file")
    p.add_argument("--base")
    p.add_argument("--code", nargs="*")
    p.add_argument("--ignore", default="")
    p.add_argument("--no-normalize", action="store_true")
    p = sub.add_parser("verify")
    p.add_argument("expected")
    p.add_argument("actual")
    p.add_argument("--ignore", default="")
    p.add_argument("--no-normalize", action="store_true")
    p = sub.add_parser("record")
    p.add_argument("--prov", required=True)
    p = sub.add_parser("overrides")
    p.add_argument("diff")
    p.add_argument("--out", required=True)
    p.add_argument("--kinds", default=",".join(ov.SUPPORTED))
    a = ap.parse_args(argv)
    if a.cmd == "record":
        if not rest:
            raise SystemExit("record: '--' 뒤에 생성 스크립트를 주세요")
        code, data = trace.run(os.path.abspath(a.prov), rest)
        print("기록: %s · 저장한 통합문서 %d · 코드 위치 %d" % (a.prov, len(data["files"]), len(data["sites"])))
        for f in data["files"]:
            print("  " + f)
        return code
    return {"snap": cmd_snap, "status": cmd_status, "diff": cmd_diff, "verify": cmd_verify, "overrides": cmd_overrides}[a.cmd](a) or 0


if __name__ == "__main__":
    sys.exit(main())
