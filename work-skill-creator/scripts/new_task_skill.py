"""업무 스킬 뼈대 만들기 - work-skill-creator 의 뼈대(assets/task-skill-template)를 복사하고 이름을 채운다.

만드는 것
  <위치>/<이름>/SKILL.md                     입력 · 코드가 하는 일 · 판단 · 멈춤 · 검증 · 맞추며 고친 것 · 한계 (TODO 를 채운다)
  <위치>/<이름>/scripts/build.py             계획 → --apply → 다시 계산 → 검산 → --expect 재현 시험 뼈대
  <위치>/<이름>/scripts/wsc_xl.py            공용 도우미(복사본 - 스킬 폴더만 옮겨도 돌아가게. --helpers 를 주면 복사하지 않음)
  <위치>/<이름>/scripts/compare_xlsx.py      칸 비교(재현 시험용 복사본)
  <위치>/<이름>/references/procedure.md      사람이 하던 작업 절차
  <위치>/<이름>/references/lessons.md        재발 방지 체크리스트
  <위치>/<이름>/references/known_diffs.txt   재현 시험에서 뺄 범위(이유가 적힌 의도적 차이)
  <위치>/<이름>/references/judgment_example.json  판단 입력 형식
  (--docs 를 주면 scripts/compare_docs.py 도 복사 - 워드 · PPT 결과물일 때)

이미 있는 파일은 덮어쓰지 않는다. 기본은 계획만 보여 주고, --apply 를 줘야 만든다.

사용:
  python -X utf8 new_task_skill.py --name allowance-quarter --title "대손충당금 분기 결산파일" [--kind Q] [--dest .claude/skills] [--apply]
  --dest 기본 = 지금 폴더의 .claude/skills (그 작업 폴더에서만 쓰는 프로젝트 스킬)
          전역(어느 폴더에서나) = %USERPROFILE%/.claude/skills
  --helpers <우산 스킬 scripts 폴더>  한 건의 공통 도우미를 불러 쓴다(wsc_xl.py 를 복사하지 않음 - 두 벌이 생기지 않게)
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import re
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL_ROOT = HERE.parent
TEMPLATE = SKILL_ROOT / "assets" / "task-skill-template"
NAME_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
KIND_WORD = {"M": "달", "Q": "분기", "H": "반기", "Y": "연도"}
RESERVED = ("anthropic", "claude")


def check_name(name: str) -> list[str]:
    errs = []
    if not NAME_RE.match(name):
        errs.append("이름은 영문 소문자 · 숫자 · 하이픈(-)만 씁니다(예: allowance-quarter).")
    if len(name) > 64:
        errs.append("이름은 64자 이내.")
    for w in RESERVED:
        if w in name:
            errs.append(f"이름에 '{w}' 는 쓸 수 없습니다.")
    return errs


def clean_title(title: str) -> str:
    """업무 이름을 SKILL.md 머리(YAML) · build.py 문자열에 넣어도 깨지지 않게 다듬는다."""
    t = title.replace('"', "'").replace("\\", "/").replace("\n", " ").strip()
    return t.replace(": ", " - ")


def git_root(path: Path) -> Path | None:
    p = path.resolve()
    for d in [p] + list(p.parents):
        if (d / ".git").exists():
            return d
    return None


def plan_files(name: str, dest: Path, docs: bool, helpers: Path | None) -> list[tuple[Path, Path | None, bool]]:
    """[(대상, 원본, 틀 여부)] - 틀(.tmpl)은 이름을 채워서, 나머지는 그대로 복사."""
    base = dest / name
    items = [
        (base / "SKILL.md", TEMPLATE / "SKILL.md.tmpl", True),
        (base / "scripts" / "build.py", TEMPLATE / "scripts" / "build.py.tmpl", True),
        (base / "scripts" / "compare_xlsx.py", HERE / "compare_xlsx.py", False),
        (base / "references" / "procedure.md", TEMPLATE / "references" / "procedure.md.tmpl", True),
        (base / "references" / "lessons.md", TEMPLATE / "references" / "lessons.md.tmpl", True),
        (base / "references" / "known_diffs.txt", TEMPLATE / "references" / "known_diffs.txt.tmpl", True),
        (base / "references" / "judgment_example.json", TEMPLATE / "references" / "judgment_example.json.tmpl", True),
    ]
    if helpers is None:
        items.insert(2, (base / "scripts" / "wsc_xl.py", HERE / "wsc_xl.py", False))
    if docs:
        items.append((base / "scripts" / "compare_docs.py", HERE / "compare_docs.py", False))
    return items


def fill(text: str, values: dict) -> str:
    for k, v in values.items():
        text = text.replace("{{" + k + "}}", v)
    return text


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="업무 스킬 뼈대 만들기")
    ap.add_argument("--name", required=True, help="스킬 이름(영문 소문자 · 하이픈)")
    ap.add_argument("--title", required=True, help="업무 이름(한글 가능) - 예: 대손충당금 분기 결산파일")
    ap.add_argument("--kind", default="Q", choices=list(KIND_WORD), help="기간 단위: M 월 · Q 분기 · H 반기 · Y 연")
    ap.add_argument("--dest", default=".claude/skills", help="스킬을 둘 폴더(기본: 지금 폴더의 .claude/skills)")
    ap.add_argument("--docs", action="store_true", help="워드 · PPT 비교 도구도 넣기")
    ap.add_argument("--helpers", default="", help="한 건의 공통 도우미 폴더(우산 스킬 scripts) - wsc_xl.py 를 복사하지 않고 불러 씀")
    ap.add_argument("--apply", action="store_true", help="실제로 만든다(없으면 계획만)")
    a = ap.parse_args(argv)

    errs = check_name(a.name)
    if errs:
        for e in errs:
            print(f"[멈춤] {e}")
        return 1
    dest = Path(a.dest).expanduser()
    helpers = Path(a.helpers).expanduser().resolve() if a.helpers else None
    if helpers is not None and not helpers.is_dir():
        print(f"[멈춤] 공통 도우미 폴더가 없습니다: {helpers}")
        return 1
    title = clean_title(a.title)
    scripts_dir = (dest / a.name / "scripts").resolve()
    if helpers is not None:
        rel = os.path.relpath(helpers, scripts_dir).replace("\\", "/")
        helpers_expr = f'(HERE / "{rel}").resolve()'
    else:
        helpers_expr = "None"
    values = {"NAME": a.name, "TITLE": title, "KIND": a.kind, "PERIOD": KIND_WORD[a.kind],
              "DATE": dt.date.today().isoformat(), "HELPERS": helpers_expr}
    items = plan_files(a.name, dest, a.docs, helpers)
    missing_src = [s for _, s, _ in items if s is not None and not s.exists()]
    if missing_src:
        for s in missing_src:
            print(f"[멈춤] 뼈대 파일이 없습니다: {s}")
        return 1

    print(f"[{'만들기' if a.apply else '계획'}] 업무 스킬 '{a.name}' ({title}, 기간 단위: {KIND_WORD[a.kind]})")
    print(f"  위치: {(dest / a.name).resolve()}")
    if title != a.title:
        print(f"  업무 이름을 '{title}' 로 다듬었습니다(따옴표 · ': ' 는 SKILL.md 머리를 깨뜨림).")
    if helpers is not None:
        print(f"  공통 도우미: {helpers} (wsc_xl.py 는 복사하지 않음)")
        if not (helpers / "wsc_xl.py").exists():
            print(f"  [주의] 공통 도우미 폴더에 wsc_xl.py 가 없습니다 - 만들면서 그 폴더에 한 벌 복사합니다.")
    g = git_root(dest)
    if g is not None:
        print(f"  [경고] 만들 위치가 git 저장소({g}) 안입니다. 고객 정보(회사명 · 계정 · 법인)가 들어갈 스킬이면 "
              f"저장소 밖 업무(건) 폴더에 만드세요.")
    exists = [t for t, _, _ in items if t.exists()]
    for t, s, tmpl in items:
        mark = "있음 - 그대로 둠" if t.exists() else ("뼈대 채움" if tmpl else "복사")
        print(f"  {'·' if t.exists() else '+'} {t.relative_to(dest)}  ({mark})")
    if not a.apply:
        print("\n계획만 보여 줬습니다. 이름 · 위치를 확인받은 뒤 --apply 로 만듭니다.")
        return 0
    made = 0
    if helpers is not None and not (helpers / "wsc_xl.py").exists():
        shutil.copy2(HERE / "wsc_xl.py", helpers / "wsc_xl.py")
    for t, s, tmpl in items:
        if t.exists():
            continue
        t.parent.mkdir(parents=True, exist_ok=True)
        if tmpl:
            t.write_text(fill(s.read_text(encoding="utf-8"), values), encoding="utf-8")
        else:
            shutil.copy2(s, t)
        made += 1
    print(f"\n만듦 {made}개" + (f" · 이미 있어서 그대로 둔 것 {len(exists)}개" if exists else ""))
    print("다음: SKILL.md · scripts/build.py 의 TODO 를 업무에 맞게 채우고, check_skill.py 로 점검합니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
