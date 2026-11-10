# -*- coding: utf-8 -*-
"""PreToolUse·PostToolUse 훅 - 도구 호출이 사람이 고친 Office 문서를 덮어쓰지 못하게 한다.

맡는 범위(파이썬 저장은 office-guard/runtime/sitecustomize.py 가 맡는다):
  막기  : Write·Edit, cp·mv·copy·Copy-Item·Move-Item, 리다이렉트, Set-Content·Out-File·tee, COM .SaveAs("경로")
  확인  : office_guard.py accept (사용자 승인 창)
  백업  : 쓰기 가능성이 있는 명령에 보이는 파일, COM 을 쓰는 스크립트에 적힌 파일
  기록  : 명령 전에 코드 소유였던 파일이 명령 동안 바뀌면 바뀐 내용도 코드 소유로 친다(별도 COM 후처리).
          사람 파일을 복사한 사본은 그 경로에서만 코드 소유로 치고, 옮긴 사람 파일은 계속 사람 파일로 둔다
  사본  : 판 이름(_v2, _전, _백업 등)을 붙인 새 빌더와 코드가 만든 문서의 판 이름 사본은 만들지 못하게 한다
읽기만 하는 명령은 건드리지 않는다. 훅 내부 오류는 작업을 막지 않는다.
"""
import json
import os
import re
import sys

sys.path.insert(0, os.environ.get("OFFICE_GUARD_HOME") or os.path.join(os.path.expanduser("~"), ".claude", "skills", "office-guard"))
import office_guard as og  # noqa: E402

EXT = r"(?:xlsx|xlsm|docx|pptx|dsd)"
PATH_RE = re.compile(r"""([A-Za-z]:[\\/][^"'\n\r<>|*?]+?\.%s|[^\s"'\n\r<>|*?=,()]+\.%s)\b""" % (EXT, EXT), re.I)
SAVEAS_RE = re.compile(r"""\.(?:SaveAs2?|SaveCopyAs)\(\s*r?["']([^"'{}$\n]+\.%s)["']""" % EXT, re.I)
PY_SCRIPT_RE = re.compile(r"""([A-Za-z]:[\\/][^"'\n\r<>|*?]+?\.py|[^\s"'\n\r<>|*?=,()]+\.py)\b""")
COM_RE = re.compile(r"win32com|Dispatch(?:Ex)?\(|New-Object\s+-ComObject|xlwings", re.I)

COPY_CMDS = {"cp", "copy", "mv", "move", "copy-item", "move-item", "cpi", "mi"}
MOVE_CMDS = {"mv", "move", "move-item", "mi"}
WRITE_CMDS = {"set-content", "add-content", "out-file", "tee", "tee-object"}
READ_ONLY = {"ls", "dir", "cat", "head", "tail", "grep", "rg", "find", "wc", "stat", "file", "echo", "pwd", "which", "where",
             "git", "gh", "test", "get-childitem", "gci", "get-item", "gi", "get-content", "gc", "test-path", "select-string",
             "write-output", "write-host", "resolve-path", "get-filehash", "measure-object", "select-object", "sha1sum", "md5sum",
             "sort", "uniq", "sed", "awk", "cut", "xargs", "cd", "set-location"}


def strip_heredocs(s):
    s = re.sub(r"""<<-?\s*(['"]?)(\w+)\1[^\n]*\n.*?\n\s*\2\b""", " ", s, flags=re.S)
    return re.sub(r"""@(['"])\r?\n.*?\r?\n\1@""", " ", s, flags=re.S)


def segments(s):
    s = re.sub(r"\d?>&\d", " ", s)
    out, cur, q = [], [], None
    for c in s:
        if q:
            q = None if c == q else q
            cur.append(c)
        elif c in "\"'":
            q = c
            cur.append(c)
        elif c in ";\n&|":
            out.append("".join(cur))
            cur = []
        else:
            cur.append(c)
    out.append("".join(cur))
    return [x.strip() for x in out if x.strip()]


def raw_tokens(seg):
    return re.findall(r'"[^"]*"|\'[^\']*\'|\S+', seg)


def unq(t):
    return t[1:-1] if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'" else t


def cmd_name(tok):
    return re.sub(r"\.exe$", "", os.path.basename(unq(tok).replace("\\", "/")).lower())


def resolve(p, cwd):
    p = unq(p).strip()
    p = re.sub(r"^\$env:USERPROFILE|^\$HOME|^%USERPROFILE%", lambda m: os.path.expanduser("~"), p, flags=re.I)
    if p.startswith("~"):
        p = os.path.expanduser(p)
    m = re.match(r"^/([a-zA-Z])/(.*)$", p)
    if m:
        p = m.group(1).upper() + ":/" + m.group(2)
    if not os.path.isabs(p):
        p = os.path.join(cwd, p)
    return os.path.normpath(p)


def _is_flag(t):
    return t.startswith("-") or bool(re.match(r"^/[a-zA-Z]$", t))


def _named(tokens, names):
    low = [t.lower() for t in tokens]
    for n in names:
        if n in low:
            i = low.index(n)
            if i + 1 < len(tokens):
                return i + 1
    return None


def _positional(tokens):
    out, i, value_flags = [], 0, {"-path", "-literalpath", "-destination", "-filepath", "-value", "-encoding", "-t"}
    while i < len(tokens):
        t = unq(tokens[i])
        if _is_flag(t):
            i += 2 if t.lower() in value_flags else 1
            continue
        out.append(t)
        i += 1
    return out


def write_targets(command, cwd, moves=None):
    """이 명령이 쓰는 파일 경로 목록(종류 무관). moves 를 주면 복사·이동 대상 -> [종류, 원천] 을 채운다."""
    found = []
    for m in SAVEAS_RE.finditer(command):
        found.append(resolve(m.group(1), cwd))
    for seg in segments(strip_heredocs(command)):
        toks = raw_tokens(seg)
        if not toks:
            continue
        name, rest = cmd_name(toks[0]), toks[1:]
        if name in ("cd", "set-location", "sl", "pushd") and rest:     # cd 뒤의 상대 경로는 옮겨 간 폴더 기준
            cwd = resolve(_positional(rest)[0] if _positional(rest) else rest[-1], cwd)
            continue
        if name in COPY_CMDS:
            i = _named(rest, ["-destination", "-dest"])
            pos = _positional(rest)
            if i is not None:
                dest, srcs = unq(rest[i]), pos or [unq(rest[j + 1]) for j, t in enumerate(rest[:-1]) if t.lower() in ("-path", "-literalpath")]
            elif len(pos) >= 2:
                dest, srcs = pos[-1], pos[:-1]
            else:
                dest, srcs = None, []
            if dest:
                d = resolve(dest, cwd)
                kind = "move" if name in MOVE_CMDS else "copy"
                if os.path.isdir(d) or dest.endswith(("/", "\\")):
                    pairs = [(os.path.join(d, os.path.basename(unq(s).replace("\\", "/"))), s) for s in srcs]
                else:
                    pairs = [(d, srcs[0] if len(srcs) == 1 else None)]
                for t, src in pairs:
                    found.append(t)
                    if moves is not None and src:
                        moves[t] = [kind, resolve(src, cwd)]
        elif name in WRITE_CMDS:
            i = _named(rest, ["-path", "-literalpath", "-filepath"])
            pos = _positional(rest)
            if i is not None:
                found.append(resolve(rest[i], cwd))
            elif pos:
                found.append(resolve(pos[0], cwd))
        elif name == "dd":
            found += [resolve(t[3:], cwd) for t in toks if unq(t).startswith("of=")]
        for k, t in enumerate(toks):
            if t[0] in "\"'":
                continue
            m = re.match(r"^\d?>>?(.*)$", t)
            if m:
                tgt = m.group(1) or (unq(toks[k + 1]) if k + 1 < len(toks) else "")
                if tgt:
                    found.append(resolve(tgt, cwd))
    return list(dict.fromkeys(found))


def overwrite_targets(command, cwd):
    """이 명령이 덮어쓰는 Office 파일 경로 목록."""
    return [p for p in write_targets(command, cwd) if og.is_guarded(p)]


VERSION_RE = re.compile(r"^(?P<base>.+?)(?:_v\d+.*|[_ -]?(?:수정전|수정후|백업|전|후|bak|old|backup)\d*)$", re.I)
DOC_VERSION_RE = re.compile(r"^(?P<base>.+?)[_ -]+(?:v\d+.*|(?:수정전|수정후|백업|사본|복사본|전|후|bak|old|backup|copy)\d*)$", re.I)
BUILDER_RE = re.compile(r"openpyxl|xlsxwriter|win32com|xlwings|docx|pptx|xlsx|xlsm|\.dsd\b", re.I)


def _office_code(path):
    try:
        if os.path.getsize(path) > 2_000_000:
            return False
        with open(path, encoding="utf-8", errors="ignore") as f:
            return bool(BUILDER_RE.search(f.read()))
    except OSError:
        return False


def version_copy(path, content=None):
    """판 이름을 붙인 새 빌더나 코드가 만든 문서의 판 이름 사본이면 안내문, 아니면 None."""
    d, name = os.path.split(path)
    stem, ext = os.path.splitext(name)
    if os.path.exists(path) or og.exempt_place(path):
        return None
    try:
        sibs = os.listdir(d)
    except OSError:
        return None
    if ext.lower() == ".py":
        m = VERSION_RE.match(stem)
        if not m:
            return None
        base = m.group("base")
        same = [x for x in sibs if x.lower().endswith(".py") and (x[:-3] == base or x[:-3].startswith(base + "_v"))]
        if not same:
            return None
        built = name.lower().startswith("_build") or any(x.lower().startswith("_build") for x in os.path.normpath(d).split(os.sep))
        if not built and not (content and BUILDER_RE.search(content)) and not any(_office_code(os.path.join(d, x)) for x in same):
            return None                        # Office 문서를 만드는 빌더가 아니면 관여하지 않는다
        return ("[사본 금지] 판 이름을 붙인 빌더를 새로 만들지 않습니다: %s\n"
                "  같은 이름의 빌더를 고치세요. 문서를 저장할 때 코드가 _history 에 자동으로 남습니다.\n"
                "  옛 판 보기: python -X utf8 ~/.claude/skills/office-guard/builder_history.py list <빌더.py>" % name)
    if not og.is_guarded(path):
        return None
    m = DOC_VERSION_RE.match(stem)
    if not m:
        return None
    base, last = m.group("base").rstrip(" _-"), None
    for x in sibs:
        xs, xe = os.path.splitext(x)
        if xe.lower() != ext.lower() or x.startswith("~$"):
            continue
        xm = DOC_VERSION_RE.match(xs)
        if xs != base and not (xm and xm.group("base").rstrip(" _-") == base):
            continue
        q = os.path.join(d, x)
        try:
            if last is None or os.path.getmtime(q) > os.path.getmtime(last):
                last = q
        except OSError:
            pass
    if not last:
        return None
    try:
        v = og.judge(last)
        if not (v["ok"] and v["own"]) or og.is_own_at(og.sha1(last), last):
            return None                        # 앞 판이 사람 파일이면 새 판을 따로 만드는 편이 안전하다
    except OSError:
        return None
    return ("[사본 금지] 코드가 만든 문서에 판 이름을 붙인 새 사본을 만들지 않습니다: %s\n"
            "  앞 판 %s 은 코드가 만든 그대로입니다. 그 파일을 고치세요(고치기 전 내용은 _history, _backup 에 남습니다).\n"
            "  앞 판을 그대로 남기고 새 판을 내야 하면(사용자가 앞 판을 마감했거나 정본이라고 한 경우) 먼저 정본 전환을 기록하세요:\n"
            "    python -X utf8 %s handover \"%s\"" % (name, os.path.basename(last), og.TOOL, last))


def visible_paths(text, cwd):
    out = []
    for m in PATH_RE.finditer(text.replace("\\\\", "\\")):
        p = resolve(m.group(1), cwd)
        if og.is_guarded(p) and p not in out:
            out.append(p)
    return out


def read_only(command):
    segs = segments(strip_heredocs(command))
    return bool(segs) and all(raw_tokens(s) and cmd_name(raw_tokens(s)[0]) in READ_ONLY and not re.search(r"(^|\s)\d?>", s) for s in segs)


def script_paths(command, cwd):
    """명령이 실행하는 .py 안에 적힌 문서 경로와 그 스크립트가 COM 을 쓰는지."""
    paths, com = [], False
    for m in PY_SCRIPT_RE.finditer(command.replace("\\\\", "\\")):
        p = resolve(m.group(1), cwd)
        try:
            if not os.path.isfile(p) or os.path.getsize(p) > 2_000_000:
                continue
            with open(p, encoding="utf-8", errors="ignore") as f:
                src = f.read()
        except OSError:
            continue
        com = com or bool(COM_RE.search(src))
        paths += [x for x in visible_paths(src, os.path.dirname(p)) if x not in paths]
    return paths, com


def pending_file(tool_use_id):
    return os.path.join(og.state_dir(), "pending", re.sub(r"\W", "_", tool_use_id or "none") + ".json")


def emit(decision, reason):
    print(json.dumps({"hookSpecificOutput": {"hookEventName": "PreToolUse", "permissionDecision": decision,
                                             "permissionDecisionReason": reason}}, ensure_ascii=False))


def pre(data):
    name, ti = data.get("tool_name") or "", data.get("tool_input") or {}
    cwd = data.get("cwd") or os.getcwd()
    if name in ("Write", "Edit"):
        p = resolve(ti.get("file_path") or "", cwd)
        msg = version_copy(p, ti.get("content") or "") if name == "Write" else None
        if msg:
            return emit("deny", msg)
        if og.is_guarded(p):
            v = og.judge(p)
            if not v["ok"]:
                return emit("deny", og.refusal(p, v))
        return
    if name not in ("Bash", "PowerShell"):
        return
    command = ti.get("command") or ""
    if re.search(r"office_guard\.py[\"']?\s+accept\b", command):
        return emit("ask", "사람이 고친 파일을 코드가 다시 덮어써도 되는 상태로 표시합니다. 사용자가 확인한 경우에만 허용하세요.")
    if re.search(r"office_guard\.py[\"']?\s+handover\b.*--back\b", command, re.S):
        return emit("ask", "정본을 사람 파일에서 코드로 되돌립니다. 그 뒤로 빌더가 이 문서를 다시 만들 수 있습니다. 사용자가 확인한 경우에만 허용하세요.")
    moves = {}
    targets = write_targets(command, cwd, moves)
    refused = [m for m in (version_copy(p) for p in targets) if m]
    for p in targets:
        if og.is_guarded(p):
            v = og.judge(p)
            if not v["ok"]:
                refused.append(og.refusal(p, v))
    if refused:
        return emit("deny", "\n".join(refused))
    if read_only(command):
        return
    shown = visible_paths(command, cwd)
    literal, com = script_paths(command, cwd)
    guarded = [p for p in targets if og.is_guarded(p)]
    made = [m for m in (version_copy(p) for p in dict.fromkeys(shown + literal) if p not in targets) if m]
    if made:
        return emit("deny", "\n".join(made))
    watch, human = {}, {}
    for p in dict.fromkeys(shown + literal + guarded):
        if not os.path.isfile(p):
            if p in shown or p in guarded:
                watch[p] = None
            continue
        if og.exempt_place(p):
            continue
        v = og.judge(p)
        if v["own"]:
            watch[p] = og.sha1(p)
            continue
        human.setdefault(og.sha1(p), []).append(p)      # 사람 문서: 복사·이동돼도 코드 소유로 올리지 않게 내용을 적어 둔다
        if p in shown or com:
            try:
                og.backup(p)
            except OSError:
                pass
    if watch:
        f = pending_file(data.get("tool_use_id"))
        os.makedirs(os.path.dirname(f), exist_ok=True)
        with open(f, "w", encoding="utf-8") as fh:
            json.dump({"watch": watch, "human": human, "moves": {p: m for p, m in moves.items() if p in watch}}, fh, ensure_ascii=False)


def _maybe_doc(src):
    """복사·이동의 원천이 문서일 수 있는지 - 변수나 와일드카드처럼 이름을 알 수 없으면 문서로 본다."""
    name = os.path.basename(src)
    return og.is_guarded(name) or "." not in name or bool(re.search(r"[$*?%`]", name))


def post(data):
    f = pending_file(data.get("tool_use_id"))
    if not os.path.exists(f):
        return
    with open(f, encoding="utf-8") as fh:
        saved = json.load(fh)
    os.remove(f)
    if not isinstance(saved.get("watch"), dict):        # 옛 형식: 경로 -> 해시
        saved = {"watch": saved}
    human, moves = saved.get("human") or {}, saved.get("moves") or {}
    for p, before in saved["watch"].items():
        if not os.path.isfile(p):
            continue
        cur = og.sha1(p)
        if cur == before or og.is_own(cur):
            continue
        how = moves.get(p)
        if cur in human:                                # 명령 전에 사람 문서였던 내용이 이 경로에 생겼다
            if any(not og.same_path(x, p) and os.path.isfile(x) and og.sha1(x) == cur for x in human[cur]):
                og.remember_at(cur, p)                  # 복사: 이 경로에서만 코드 소유
        elif how and _maybe_doc(how[1]):                # 원천을 미리 보지 못한 문서 복사·이동
            if how[0] == "copy":
                og.remember_at(cur, p)
        else:
            og.remember(cur)


def main():
    try:
        data = json.load(sys.stdin)
    except Exception:
        return
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        (post if data.get("hook_event_name") == "PostToolUse" else pre)(data)
    except Exception:
        pass


if __name__ == "__main__":
    main()
    sys.exit(0)
