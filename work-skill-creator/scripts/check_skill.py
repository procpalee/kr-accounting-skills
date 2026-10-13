"""업무 스킬 점검 - 스킬을 만들거나 고친 뒤, 넘기기 전에 돌린다.

보는 것
  SKILL.md 머리(name · description) · 이름 규칙 · 설명 길이(1,024자) · 요청 문구
  업무 스킬의 필수 절(입력 · 코드가 하는 일 · 판단 · 멈춤 · 검증 · 맞추며 고친 것 · 한계) - 절 이름만 보고 순서는 보지 않는다
  남은 TODO(붙여 쓴 것까지) · 뼈대 자리표시 / SKILL.md 가 가리키는 파일이 있는지 / 파이썬 문법
  비밀값처럼 보이는 글자(API 키 · 토큰 · 비밀번호) · 비밀값 파일(.env · *.key · …키.txt) / 내 PC 절대 경로
  --forbid 목록의 단어(고객사명 등) - 글 파일과 엑셀 · 워드 · PPT 안 글자까지

사용: python -X utf8 check_skill.py <스킬 폴더> [--kind task|any] [--forbid 단어목록.txt] [--strict]
  --kind any  : 업무 스킬 필수 절 검사를 건너뜀(일반 스킬)
  --strict    : 남은 TODO 도 오류로 본다(넘기기 직전)
종료 코드: 오류 없음 0 · 오류 있음 1
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

NAME_RE = re.compile(r"^[a-z0-9]+(-[a-z0-9]+)*$")
SECTIONS = {  # 절 이름에 이 낱말 중 하나가 있어야 한다
    "입력": ("입력",),
    "코드가 하는 일": ("코드",),
    "판단(사용자에게 묻기)": ("판단",),
    "멈춤 · 경고": ("멈춤", "중단", "경고"),
    "산출 · 검증": ("검증", "확인"),
    "맞추며 고친 것(재현 기록)": ("맞추", "재현", "고친 것"),
    "한계 · 가정": ("한계", "가정"),
}
SECRET_PATTERNS = [
    (re.compile(r"(?i)(service_?key|api_?key|secret|token|password|passwd|비밀번호)\s*[:=]\s*['\"]?[A-Za-z0-9%+/=_\-]{16,}"), "키 · 토큰 · 비밀번호 값"),
    (re.compile(r"\bsk-[A-Za-z0-9_\-]{20,}"), "API 키(sk-…)"),
    (re.compile(r"\bAKIA[0-9A-Z]{16}\b"), "AWS 키"),
    (re.compile(r"\bgh[pousr]_[A-Za-z0-9]{30,}\b"), "GitHub 토큰"),
    (re.compile(r"\b[A-Fa-f0-9]{48,}\b"), "긴 16진수(키일 수 있음)"),
]
USER_PATH = re.compile(r"(?i)[A-Z]:[\\/]+Users[\\/]+(?!<|%|\{|\$|사용자|PC>|내 이름)[^\\/\s'\"`]+")
TEXT_EXT = {".md", ".py", ".json", ".yaml", ".yml", ".txt", ".tmpl", ".cfg", ".ini", ".toml", ".ps1", ".bat", ".cmd",
            ".csv", ".tsv", ".xml", ".html", ".htm", ".js", ".sql"}
OFFICE_EXT = {".xlsx", ".xlsm", ".docx", ".pptx"}
# 스킬 폴더에 두면 안 되는 비밀값 파일 이름
SECRET_FILE = re.compile(r"(?i)(^\.env$|\.env\.|\.key$|\.pem$|\.pfx$|secret|token|credential|api_?key|api키|비밀번호|인증키)")
PLACEHOLDER = re.compile(r"\{\{(NAME|TITLE|KIND|PERIOD|DATE|HELPERS)\}\}")


def office_text(path: Path) -> str:
    """엑셀 · 워드 · PPT 안의 글자(XML) - 금지어(고객사명 등)를 찾는 데 쓴다."""
    import html
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            raw = "\n".join(z.read(n).decode("utf-8", "ignore") for n in z.namelist() if n.endswith((".xml", ".rels")))
        return html.unescape(raw)  # openpyxl 은 한글을 &#44032; 같은 문자 참조로 쓴다
    except Exception:
        return ""


def read_frontmatter(text: str) -> tuple[dict, str, list[str]]:
    errs = []
    if not text.startswith("---"):
        return {}, text, ["SKILL.md 가 '---' 머리(frontmatter)로 시작하지 않습니다."]
    end = text.find("\n---", 3)
    if end < 0:
        return {}, text, ["머리(frontmatter)를 닫는 '---' 가 없습니다."]
    head, body = text[3:end].strip("\n"), text[end + 4:]
    meta = {}
    try:
        import yaml  # 있으면 정확히
        meta = yaml.safe_load(head) or {}
        if not isinstance(meta, dict):
            errs.append("머리(frontmatter)가 '키: 값' 모양이 아닙니다.")
            meta = {}
    except ImportError:
        for line in head.splitlines():
            m = re.match(r"^([A-Za-z_][\w-]*):\s*(.*)$", line)
            if m:
                meta[m.group(1)] = m.group(2).strip().strip("'\"")
    except Exception as e:
        errs.append(f"머리(frontmatter)를 읽지 못했습니다: {e}")
    return meta, body, errs


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="업무 스킬 점검")
    ap.add_argument("skill_dir")
    ap.add_argument("--kind", default="task", choices=["task", "any"])
    ap.add_argument("--forbid", default="", help="들어가면 안 되는 단어 목록 파일(한 줄에 하나)")
    ap.add_argument("--strict", action="store_true")
    a = ap.parse_args(argv)

    root = Path(a.skill_dir)
    errors, warns, oks = [], [], []
    sk = root / "SKILL.md"
    if not sk.exists():
        print(f"[오류] SKILL.md 가 없습니다: {sk}")
        return 1
    text = sk.read_text(encoding="utf-8")
    meta, body, errs = read_frontmatter(text)
    errors += errs

    name = str(meta.get("name") or "")
    desc = str(meta.get("description") or "")
    if not name:
        errors.append("name 이 없습니다.")
    elif not NAME_RE.match(name) or len(name) > 64:
        errors.append(f"name '{name}' - 영문 소문자 · 숫자 · 하이픈만, 64자 이내.")
    elif any(w in name for w in ("anthropic", "claude")):
        errors.append(f"name '{name}' 에 anthropic · claude 는 쓸 수 없습니다.")
    else:
        oks.append(f"name: {name}")
        if root.resolve().name != name:
            warns.append(f"폴더 이름({root.resolve().name})과 name({name})이 다릅니다 - 같게 두는 것이 안전합니다.")
    if not desc:
        errors.append("description 이 없습니다 - 무엇을 하는지 + 어떤 요청에 쓰는지를 적습니다.")
    else:
        if len(desc) > 1024:
            errors.append(f"description 이 {len(desc)}자 - 1,024자 이내로 줄입니다.")
        else:
            oks.append(f"description {len(desc)}자")
        if not re.search(r"[\"“'‘]", desc):
            warns.append("description 에 사용자가 실제로 쓰는 요청 문구(따옴표)가 없습니다 - 스킬이 잘 안 불릴 수 있습니다.")
        if "TODO" in desc:
            (errors if a.strict else warns).append("description 에 TODO 가 남아 있습니다.")

    lines = text.count("\n") + 1
    if lines > 500:
        warns.append(f"SKILL.md 가 {lines}줄 - 500줄 안쪽으로 줄이고 자세한 내용은 references/ 로 옮깁니다.")

    if a.kind == "task":
        heads = [h.strip() for h in re.findall(r"^#{2,3}\s+(.+)$", body, re.M)]
        for label, words in SECTIONS.items():
            if not any(any(w in h for w in words) for h in heads):
                errors.append(f"업무 스킬 필수 절이 없습니다: {label}")
        if not re.search(r"--apply", text):
            warns.append("SKILL.md 에 '계획 → --apply' 흐름 안내가 없습니다.")
        build = root / "scripts" / "build.py"
        if build.exists() and "--apply" not in build.read_text(encoding="utf-8", errors="ignore"):
            warns.append("scripts/build.py 에 --apply(계획 먼저) 옵션이 없습니다.")

    # 링크로 건 파일은 꼭 있어야 한다. 글 속에 적은 경로는 업무 스킬일 때만 본다(다른 스킬의 파일을 말할 수도 있어서).
    for ref in sorted(set(re.findall(r"\]\(((?:scripts|references|assets)/[^)\s]+)\)", text))):
        if not (root / ref).exists():
            errors.append(f"SKILL.md 링크가 가리키는 {ref} 가 없습니다.")
    if a.kind == "task":
        for ref in sorted(set(re.findall(r"scripts/([\w\-]+\.py)", text))):
            if not (root / "scripts" / ref).exists():
                errors.append(f"SKILL.md 가 가리키는 scripts/{ref} 가 없습니다.")
        for ref in sorted(set(re.findall(r"references/([\w\-.]+\.(?:md|json|yaml|txt))", text))):
            if not (root / "references" / ref).exists():
                warns.append(f"SKILL.md 가 가리키는 references/{ref} 가 없습니다.")

    forbid = []
    if a.forbid:  # 메모장 저장(BOM) 도 읽히게 utf-8-sig
        forbid = [w.strip() for w in Path(a.forbid).read_text(encoding="utf-8-sig").splitlines() if w.strip() and not w.startswith("#")]
    # 뼈대에서 나온 파일 - TODO · 자리표시가 남으면 안 되는 곳
    scaffolded = {"SKILL.md", "scripts/build.py", "references/procedure.md", "references/lessons.md",
                  "references/judgment_example.json", "references/known_diffs.txt"}
    todo_files = []
    for p in sorted(root.rglob("*")):
        if not p.is_file() or "__pycache__" in p.parts:
            continue
        rel = p.relative_to(root)
        rel_s = str(rel).replace("\\", "/")
        if SECRET_FILE.search(p.name):
            errors.append(f"{rel}: 비밀값 파일로 보입니다 - 키 · 비밀번호 파일은 스킬 폴더 밖(환경 변수나 사용자 폴더)에 두고 경로만 적습니다.")
        if p.suffix.lower() in OFFICE_EXT:
            if forbid:
                t = office_text(p)
                for w in forbid:
                    if w in t or w in p.name:
                        errors.append(f"{rel}: 파일 안에 들어가면 안 되는 단어 '{w}'")
            continue
        if p.suffix.lower() not in TEXT_EXT:
            continue
        t = p.read_text(encoding="utf-8", errors="ignore")
        if p.suffix == ".py":
            try:
                compile(t, str(p), "exec")
            except SyntaxError as e:
                errors.append(f"{rel}: 파이썬 문법 오류 {e.lineno}행 - {e.msg}")
        if a.kind == "task" and rel_s in scaffolded:
            lines_with_todo = [ln for ln in t.splitlines() if "TODO" in ln and not ln.lstrip().startswith("#")] \
                if rel_s == "references/known_diffs.txt" else [ln for ln in t.splitlines() if "TODO" in ln]
            if lines_with_todo:  # 'TODO키' · 'TODO_{ym}' 처럼 붙여 쓴 것도 잡는다
                todo_files.append(f"{rel}({len(lines_with_todo)}줄)")
            if PLACEHOLDER.search(t):
                errors.append(f"{rel}: 뼈대의 자리표시(NAME · TITLE 등)가 채워지지 않았습니다.")
        for rx, label in SECRET_PATTERNS:
            m = rx.search(t)
            if m:
                errors.append(f"{rel}: 비밀값처럼 보이는 글자({label}) - '{m.group(0)[:12]}…' 키는 파일 · 환경 변수에서 읽고 스킬에 적지 않습니다.")
        m = USER_PATH.search(t)
        if m:
            warns.append(f"{rel}: 내 PC 절대 경로({m.group(0)[:40]}) - 다른 PC 에서 안 돌 수 있습니다. 상대 경로 · 옵션으로 바꿉니다.")
        for w in forbid:
            if w in t or w in str(rel):
                errors.append(f"{rel}: 들어가면 안 되는 단어 '{w}'")
    if todo_files:
        (errors if a.strict else warns).append("TODO 가 남은 파일: " + ", ".join(todo_files))

    print(f"[스킬 점검] {root.resolve()}")
    for x in oks:
        print(f"  ✓ {x}")
    for x in warns:
        print(f"  [경고] {x}")
    for x in errors:
        print(f"  [오류] {x}")
    print(f"결과: 오류 {len(errors)} · 경고 {len(warns)}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())
