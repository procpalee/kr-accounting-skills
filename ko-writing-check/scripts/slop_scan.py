# -*- coding: utf-8 -*-
"""ko-writing-check 스캐너: 한국어 문서에 남은 AI 티 신호를 찾는다.

docx/pptx/xlsx/md/txt/html 본문을 문단(셀) 단위로 뽑아 references/patterns.md 의 패턴 가운데
기계로 판정할 수 있는 것만 찾는다. 나머지 판단은 사람(모델) 몫이다.
WARN 은 고칠 것, INFO 는 읽고 판단할 것. 뜻 수준 신호(M_*)는 INFO 로 세고 건수는 늘 첫 줄에 보인다.
--heads 는 제목과 그 아래 첫 단락만 뽑는다(뜻 점검·독립 검토 입력용).
규칙을 바꾸면 selftest.py 를 돌린다.

usage: python slop_scan.py FILE [FILE ...] [--info] [--prose] [--heads] [--out report.txt]
exit : 0 = WARN 없음(PASS), 1 = WARN 있음(FAIL), 2 = 입력 오류
"""
import argparse
import html
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"

HANGUL = re.compile(r"[가-힣]")
EMOJI = re.compile("[✅❌❎✨⭐\U0001F300-\U0001FAFF]")
PARTICLE = re.compile(r"""[A-Za-z0-9%)\]"'”’] (은|는|이|가|을|를|에|에서|에게|와|과|로|으로|의)(?=[\s,.)]|$)""")

# (code, level, regex, 설명). 붙임 복합어 middot(성능·품질)·화살표·※는 신호 아님.
RULES = [
    ("EMDASH", "WARN", re.compile(r"—"), "산문 속 em-dash: 콜론/괄호/문장 분리"),
    ("STAR", "WARN", re.compile(r"[★☆]"), "장식 별표: 삭제(강조는 볼드/색)"),
    ("EMOJI", "WARN", EMOJI, "장식 이모지: 삭제"),
    ("DIVIDE", "WARN", re.compile(r"÷"), "수식기호: 말로 풀기"),
    ("MIDDOT_SPACED2", "WARN", re.compile(r"[  ]·[  ].*[  ]·[  ]"), "공백 가운뎃점 나열: 쉼표 또는 /"),
    ("MIDDOT_DOUBLE", "WARN", re.compile(r" {2,}·|· {2,}"), "이중 공백 구분자: 단일 구분자"),
    ("PARTICLE_SPACE", "WARN", PARTICLE, "조사 앞 공백: 붙이기"),
    ("TL_ISSEO", "WARN", re.compile(r"에 있어서?(?=[\s,])"), "번역투 '~에 있어(서)': 생략 또는 '~에서'"),
    ("TL_TONGHAYEO", "WARN", re.compile(r"[을를] 통하여"), "번역투 '~을 통하여': '~로/~으로'"),
    ("TL_DOEEOJI", "WARN", re.compile(r"되어[지진져질짐집]"), "이중 피동: '~되다'"),
    ("TL_HALSUITDA", "WARN", re.compile(r"(이)?라고? 할 수 있"), "번역투 '~라 할 수 있다': 단정"),
    ("TL_DAEBI_SANGI", "INFO", re.compile(r"대비.{0,30}상이"), "번역투 '대비 … 상이': '~와 다름' 검토"),
    ("MIDDOT_SPACED1", "WARN", re.compile(r"\S[  ]·[  ]\S"), "가운뎃점 앞뒤 공백: 붙이기(41·47명)"),
    ("RANGE_SPACE", "INFO", re.compile(r"\d[A-Za-z%]* ~ \d"), "범위 공백: 붙이면 깔끔(1K~20K)"),
    ("ENDASH", "INFO", re.compile(r" – "), "산문 속 en-dash"),
]

# --prose: 서술형 글(리포트, 에세이, 자소서, 블로그, 칼럼). 줄 단위는 INFO, 판단은 사람이.
PROSE_RULES = [
    ("PROSE_LESSON", "INFO", re.compile(r"깨달았|깨닫게\s*되|중요성을\s*(알게|배웠|배우게|느꼈|실감)|교훈을\s*(얻|배)"),
     "교훈형 마무리: 그 뒤 실제로 바뀐 행동이나 사실로"),
    ("PROSE_NOTJUST", "INFO", re.compile(r"(?<![가-힣])(단순(한|히)?|그저|단지)\s[^.!?\n]{0,30}?아(닌|니라|니었)|에\s*그치지\s*않"),
     "'단순한 ~이 아닌'·'~에 그치지 않고' 대비 틀: 사실만 남기기"),
    ("PROSE_EMPTY_EMPH", "INFO", re.compile(r"(?<![가-힣])(정말|매우|굉장히|진정한|진정으로|무엇보다)(?=\s)"),
     "빈 강조어: 지우거나 숫자·사실로"),
    ("PROSE_CLICHE_REPORT", "INFO",
     re.compile(r"본\s*(보고서|리포트|레포트|글)에서는|알아보고자\s*한다|이상으로\s*[^.\n]{0,20}살펴보았|시사하는\s*바가\s*(크|많)"),
     "리포트 상투구: 바로 주장·사실로 시작/끝"),
    ("PROSE_CLICHE_BLOG", "INFO",
     re.compile(r"오늘은\s[^.\n]{0,30}알아보(겠|도록|려고)|도움이\s*되(셨|었)(길|으면)|시작해\s*볼까요|꿀팁"),
     "블로그 상투구: 인사 대신 본론"),
]
# 뜻 수준 신호(references/patterns.md M 분류). 줄 단위 INFO, 판단은 사람이.
# 스캔이 PASS 여도 구호·빈말이 남는 문제(2026-11-14 제안서 사례)를 겉으로 드러나는 표지로 잡는다.
MEANING_RULES = [
    ("M_ERA", "INFO", re.compile(r"(?<![가-힣])[가-힣A-Za-z]{1,10}(하는|되는|의|인)\s*시대(에|입니다|이다|가|,|$)"),
     "'~의 시대' 진단: 지우고 사실로 시작"),
    ("M_FROMTO", "INFO", re.compile(r"(?<![가-힣])[가-힣]{1,8}에서\s+[가-힣]{1,8}까지(?![가-힣])"),
     "'A에서 B까지' 범위: 실제 항목을 나열"),
    ("M_TRANSFORM", "INFO", re.compile(r"(?<![가-힣])[가-힣]{1,8}에서\s+[가-힣]{1,8}(으)?로(?=\s*($|[,.!?:]))"),
     "변환 공식 'X에서 Y로': 누가 무엇을 하는지로"),
    ("M_TRIPLE", "INFO", re.compile(r"^[^.\n]{0,40}?[가-힣]고,\s[^,\n]{1,20}[가-힣]고,\s[^,\n]{1,25}$"),
     "짧은 3단 병렬 구호: 실제 내용 문장으로"),
    ("M_PROMISE", "INFO", re.compile(r"수준에\s*맞(게|춘|는|춰)|맞춤(형|화)|눈높이에\s*맞|최적화된|한\s*단계\s*(도약|높)"),
     "검증할 수 없는 약속: 실제로 다르게 하는 것을 쓰기"),
    # '시스템 구축', '내부회계 구축'처럼 실제 행위인 낱말(구축·확립)은 넣지 않는다(저자 기고문 오탐, 2026-11-14)
    ("M_POLICY", "INFO", re.compile(r"역량을?\s*강화|(?<![가-힣])(제고|도모|극대화|함양|고도화|내재화)(하|합|할|해|됩|되|를|와|에|,|\s|$)"),
     "정책동사: 실제 결과·행동으로"),
    ("M_CONCLUDE", "INFO", re.compile(r"결국[^.\n]{0,40}(이유|때문|이어지|연결)|(하는|한|된)\s*이유(다|입니다)[.\s]*$|(할|될)\s*(때|시점|순간)입니다"),
     "결론 공식: 앞 내용을 다시 말하지 말고 지우거나 다음 할 일로"),
    ("M_SIGNIF", "INFO", re.compile(r"주목할\s*만|게임\s*체인저|패러다임|새로운\s*지평|혁신적(인|으로)|무궁무진|획기적"),
     "의의 과장: 구체 결과나 숫자로"),
    ("M_LIST_INTRO", "INFO", re.compile(r"(크게|다음과\s*같이|아래와\s*같이)\s*(두|세|네|다섯)\s*가지로\s*(나눌|정리|요약)"),
     "열거 도입구: 지우고 바로 본론"),
    ("M_PERSONIFY", "INFO", re.compile(r"(AI|기술|시대|데이터|숫자|시장)(가|이|는)\s*(묻|말해|요구하|부르|보여주)"),
     "의인화 추상 주어: 사람·기관을 주어로"),
    # 아래는 references/patterns.md 의 패턴 가운데 업무 문서 오탐이 적은 것만(M-02·05·15·17·18·21, V-02·14)
    ("M_CLICHE", "INFO", re.compile(r"시대가\s?도래|급변하는\s|그\s?어느\s?때보다|초석|교두보|청사진|신호탄|지평을\s?(열|연)"
                                    r"|귀추가\s?주목|밝은\s?(미래|전망)|무한한\s?가능성|새로운\s?도약|간과할\s?수\s?없|한\s?획을"
                                    r"|자리매김|발돋움|녹여\s?(내|낸|냈)|선사(하|합|한|했)|향연"),
     "상투어·과장어: 지우거나 구체 사실로"),
    ("M_HEARSAY", "INFO", re.compile(r"(다는|라는)\s?(분석|평가|관측)(이다|입니다)|전문가들은|업계\s?관계자에\s?따르면|일각에서는"),
     "주체 없는 판정·출처 없는 권위: 출처를 밝히거나 직접 서술"),
    ("M_CHATBOT", "INFO", re.compile(r"좋은\s?질문|훌륭한\s?(질문|지적)|짚어\s?주셨|정리해\s?드리겠습니다"
                                     r"|정보(가|는)\s?(제한적|부족)하|최신\s?정보와\s?다를"),
     "챗봇 말투·면피: 지우고, 없는 정보는 무엇이 없는지 밝히기"),
]
# 번역투 확장(patterns.md T-03·T-04): 저자 글에도 드물게 있어 INFO
TL_EXTRA = [
    ("TL_BOYEOJI", "INFO", re.compile(r"보여[지진졌질짐집]"), "이중 피동 '보여지다': '보이다'"),
    ("TL_BOLSUITDA", "INFO", re.compile(r"(이)?라고\s?볼\s?수\s?있"), "'~라고 볼 수 있다': 단정하거나 근거를"),
]
# 수사 질문("그렇다면 ~일까요?")과 연결어미 뒤 쉼표는 저자 기고문에도 1,000자당 몇 번씩 나와 신호로 세지 않는다
# (2026-11-14 택스넷 1~5회차 시험). 'A가 아니라 B'도 저자 글에 회당 2~4번 있어 밀도가 높을 때만 센다.
NOT_BUT = re.compile(r"(이|가)?\s*아니라\s|것이\s*아니라|것은\s*아니")
NOT_BUT_MIN, NOT_BUT_PER_CHARS, MEANING_REVIEW_MIN = 5, 1000, 3
# '핵심은 ~' 분열문도 저자 블로그 글에 한두 번씩 있어 한 문서 3회 이상일 때만 센다.
CLEFT = re.compile(r"(?<![가-힣])(핵심|관건|열쇠|본질)(은|이)\s")
CLEFT_MIN = 3

QUOTE_OPEN = re.compile(r"^\s*(>\s*)?[“\"「『‘']|사전적\s*(의미|정의)|사전에\s*(따르면|서는)|(이)?라는\s*말(이\s*있|처럼)")
# "~라고 본다면"처럼 조건절로 이어지면 의견 종결이 아니다.
THINK = re.compile(r"([다라]고\s*(생각(한다|합니다|했다|했습니다|해요|된다|됩니다)|본다|봅니다)|"
                   r"[다라]는\s*생각이\s*(든다|듭니다|들었다|들었습니다))(?!면)")
CONNECTOR = re.compile(r"^\s*(또한|나아가|더불어|결론적으로|이처럼|이를\s*통해|따라서|그러므로|뿐만\s*아니라)[,\s]")
THINK_MIN, THINK_PER_CHARS = 3, 1000
CONNECTOR_MIN, CONNECTOR_RATIO = 3, 0.25


def is_heading(text):
    return text.lstrip().startswith("#")


LIST_ITEM = re.compile(r"^\s*([-*+]|\d+[.)])\s")


def front_matter_end(path):
    """md 머리말(--- ... ---)의 마지막 줄 번호. 없으면 0."""
    lines = Path(path).read_text(encoding="utf-8").splitlines()
    if not lines or lines[0].strip() != "---":
        return 0
    for n, line in enumerate(lines[1:], 2):
        if line.strip() == "---":
            return n
    return 0


def split_paragraphs(paras, group_lines=False, skip_until=0):
    """판정용 문단 목록. 셀 안의 빈 줄은 문단을 나누고, md는 빈 줄 없이 이어진 줄을 한 문단으로 묶는다.
    제목 줄과 목록 항목은 따로 세고, 표 셀과 머리말(skip_until 줄까지)은 문단으로 세지 않는다."""
    out, prev_n = [], None
    for where, text in paras:
        if where.endswith("(표)"):
            prev_n = None
            continue
        line_no = re.match(r"L(\d+)", where)
        if line_no and int(line_no.group(1)) <= skip_until:
            continue
        m = re.fullmatch(r"L(\d+)", where) if group_lines else None
        n = int(m.group(1)) if m else None
        alone = is_heading(text) or LIST_ITEM.match(text)
        if (n is not None and prev_n is not None and n == prev_n + 1 and not alone and out
                and not is_heading(out[-1][1]) and not LIST_ITEM.match(out[-1][1])):
            out[-1] = (out[-1][0], out[-1][1] + " " + text.strip())
        else:
            out.extend((where, part) for part in re.split(r"\n\s*\n", text) if HANGUL.search(part))
        prev_n = n
    return out


def prose_totals(paras, show_info):
    """문서 전체 기준 서술형 신호. paras = [(where, text)] (한글 있는 본문 문단)."""
    hits = []
    body = [(w, t) for w, t in paras if not is_heading(t)]
    if not body:
        return hits
    first_w, first_t = body[0]
    if show_info and QUOTE_OPEN.search(first_t):
        hits.append(("INFO", "PROSE_QUOTE_OPEN", first_w, first_t.strip()[:50], "명언·사전 정의로 시작: 장면이나 사실로"))
    chars = sum(len(t) for _w, t in body)
    think = sum(len(THINK.findall(t)) for _w, t in body)
    if think >= THINK_MIN and think * THINK_PER_CHARS >= chars:
        rate = think * THINK_PER_CHARS / chars
        hits.append(("WARN", "PROSE_THINK", "(문서 전체)", f"'~라고 생각한다' {think}회 / 1,000자당 {rate:.1f}회 ({chars}자)",
                     "사실·행동으로 말하거나 단정"))
    conn = sum(1 for _w, t in body if CONNECTOR.search(t))
    if conn >= CONNECTOR_MIN and conn >= CONNECTOR_RATIO * len(body):
        hits.append(("WARN", "PROSE_CONNECTOR", "(문서 전체)", f"연결어로 시작하는 문단 {len(body)}개 중 {conn}개",
                     "꼭 필요한 곳 하나만 남기기"))
    return hits


def is_marker(text):
    """표의 해당없음 마커(—, -, ―)만 있는 셀."""
    return text.strip() in {"—", "-", "―", "–", "N/A"}


def parent_map(root):
    return {c: p for p in root.iter() for c in p}


def paragraphs_xml(root, p_tag, t_tag):
    """t 요소를 가장 가까운 p 조상에 모은다(텍스트박스 중첩 문단도 분리)."""
    parents = parent_map(root)
    order, buf = [], {}
    for p in root.iter(p_tag):
        order.append(p)
        buf[p] = []
    for t in root.iter(t_tag):
        node = parents.get(t)
        while node is not None and node.tag != p_tag:
            node = parents.get(node)
        if node is not None and t.text:
            buf[node].append(t.text)
    return [(i, "".join(buf[p]), p, parents) for i, p in enumerate(order, 1)]


def has_ancestor(node, parents, tag):
    node = parents.get(node)
    while node is not None:
        if node.tag == tag:
            return True
        node = parents.get(node)
    return False


def rels(z, path):
    base = posixpath.dirname(path)
    rp = posixpath.join(base, "_rels", posixpath.basename(path) + ".rels")
    if rp not in z.namelist():
        return {}
    out = {}
    for r in ET.fromstring(z.read(rp)).iter(PR + "Relationship"):
        tgt = r.get("Target")
        out[r.get("Id")] = tgt.lstrip("/") if tgt.startswith("/") else posixpath.normpath(posixpath.join(base, tgt))
    return out


def extract_docx(path):
    with zipfile.ZipFile(path) as z:
        parts = ["word/document.xml"] + sorted(
            n for n in z.namelist() if re.match(r"word/(header|footer|footnotes|endnotes)\d*\.xml$", n))
        for part in parts:
            if part not in z.namelist():
                continue
            root = ET.fromstring(z.read(part))
            label = "본문" if part == "word/document.xml" else Path(part).stem
            for i, text, p, parents in paragraphs_xml(root, W + "p", W + "t"):
                where = f"{label} p{i}" + (" (표)" if has_ancestor(p, parents, W + "tc") else "")
                yield where, text


def extract_pptx(path):
    with zipfile.ZipFile(path) as z:
        prs = ET.fromstring(z.read("ppt/presentation.xml"))
        rmap = rels(z, "ppt/presentation.xml")
        slides = [rmap[s.get(R + "id")] for s in prs.iter(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}sldId")]
        for n, part in enumerate(slides, 1):
            root = ET.fromstring(z.read(part))
            for i, text, p, parents in paragraphs_xml(root, A + "p", A + "t"):
                where = f"slide{n} p{i}" + (" (표)" if has_ancestor(p, parents, A + "tc") else "")
                yield where, text


def extract_xlsx(path):
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        shared = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(S + "si"):
                phonetic = {t for rph in si.iter(S + "rPh") for t in rph.iter(S + "t")}
                shared.append("".join(t.text or "" for t in si.iter(S + "t") if t not in phonetic))
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rmap = rels(z, "xl/workbook.xml")
        for sh in wb.iter(S + "sheet"):
            part = rmap.get(sh.get(R + "id"))
            if not part or part not in names:
                continue
            name = sh.get("name")
            for c in ET.fromstring(z.read(part)).iter(S + "c"):
                t, v = c.get("t"), c.find(S + "v")
                if t == "s" and v is not None:
                    text = shared[int(v.text)]
                elif t == "inlineStr":
                    text = "".join(x.text or "" for x in c.iter(S + "t"))
                elif t == "str" and v is not None:
                    text = v.text or ""
                else:
                    continue
                yield f"{name}!{c.get('r')}", text
        for part in sorted(n for n in names if re.match(r"xl/drawings/drawing\d+\.xml$", n)):
            root = ET.fromstring(z.read(part))
            for i, text, _p, _par in paragraphs_xml(root, A + "p", A + "t"):
                yield f"{Path(part).stem}(도형) p{i}", text


def extract_md(path):
    fence = False
    for n, line in enumerate(Path(path).read_text(encoding="utf-8").splitlines(), 1):
        if line.lstrip().startswith(("```", "~~~")):
            fence = not fence
            continue
        if fence:
            continue
        # 인라인 코드는 자리표시자로: 지우면 "kit `a.zip`은"이 "kit 은"이 되어 조사 공백 오탐
        line = re.sub(r"`[^`]*`", "X", line)
        if line.strip().startswith("|"):
            for k, cell in enumerate(line.strip().strip("|").split("|"), 1):
                if not re.fullmatch(r"\s*:?-{2,}:?\s*", cell):
                    yield f"L{n} c{k} (표)", cell.strip()
        else:
            yield f"L{n}", line


def extract_html(path):
    src = Path(path).read_text(encoding="utf-8")
    src = re.sub(r"(?is)<(script|style|pre|code)\b.*?</\1>", " ", src)
    src = re.sub(r"(?i)<(br|/p|/div|/li|/h\d|/td|/th|/tr)\b[^>]*>", "\n", src)
    for n, line in enumerate(html.unescape(re.sub(r"<[^>]+>", "", src)).splitlines(), 1):
        if line.strip():
            yield f"L{n}", line


EXTRACTORS = {".docx": extract_docx, ".pptx": extract_pptx, ".xlsx": extract_xlsx, ".xlsm": extract_xlsx,
              ".md": extract_md, ".txt": extract_md, ".html": extract_html, ".htm": extract_html}


HEAD_STYLE = re.compile(r"(?i)heading|title|subtitle|제목|^\d$")


def _clean(s):
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def heads_html(path):
    src = Path(path).read_text(encoding="utf-8")
    src = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", src)
    prev_head = False
    for m in re.finditer(r"(?is)<(h[1-4]|p)\b([^>]*)>(.*?)</\1>", src):
        tag, attrs, inner = m.group(1).lower(), m.group(2), _clean(m.group(3))
        if not HANGUL.search(inner):
            continue
        line = src.count("\n", 0, m.start()) + 1
        if tag.startswith("h"):
            yield f"L{line}", "제목", inner
            prev_head = True
        elif len(inner) < 12:  # 구역 이름표·짧은 칩은 첫 단락으로 보지 않는다
            continue
        elif prev_head or re.search(r'class="[^"]*\b(sub|lead)\b', attrs):
            yield f"L{line}", "첫 단락", inner
            prev_head = False


def heads_md(path):
    prev_head, fence = False, False
    for n, line in enumerate(Path(path).read_text(encoding="utf-8").splitlines(), 1):
        if line.lstrip().startswith(("```", "~~~")):
            fence = not fence
        if fence or not line.strip() or not HANGUL.search(line):
            continue
        if line.lstrip().startswith("#"):
            yield f"L{n}", "제목", line.strip("# ").strip()
            prev_head = True
        elif prev_head:
            yield f"L{n}", "첫 단락", line.strip()
            prev_head = False


def heads_docx(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    prev_head = False
    for i, text, p, parents in paragraphs_xml(root, W + "p", W + "t"):
        if not text.strip() or not HANGUL.search(text) or has_ancestor(p, parents, W + "tc"):
            continue
        st = p.find(f"{W}pPr/{W}pStyle")
        if st is not None and HEAD_STYLE.search(st.get(W + "val", "")):
            yield f"p{i}", "제목", text.strip()
            prev_head = True
        elif prev_head:
            yield f"p{i}", "첫 단락", text.strip()
            prev_head = False


def heads_pptx(path):
    last, k = None, 0
    for where, text in extract_pptx(path):
        slide = where.split()[0]
        if slide != last:
            last, k = slide, 0
        if text.strip() and HANGUL.search(text) and k < 2 and not where.endswith("(표)"):
            yield where, "제목" if k == 0 else "첫 단락", text.strip()
            k += 1


HEADS = {".html": heads_html, ".htm": heads_html, ".md": heads_md, ".txt": heads_md,
         ".docx": heads_docx, ".pptx": heads_pptx}


def scan(path, show_info, prose=False):
    hits, quotes, paras = [], {"curly": 0, "straight": 0}, []
    rules = RULES + TL_EXTRA + MEANING_RULES + (PROSE_RULES if prose else [])
    not_but = chars = 0
    clefts = []
    for where, text in EXTRACTORS[path.suffix.lower()](path):
        if not text.strip() or is_marker(text):
            continue
        # 한글 없는 순수 영문/코드 셀은 대상 아님
        if not HANGUL.search(text):
            continue
        paras.append((where, text))
        quotes["curly"] += len(re.findall(r"[“”]", text))
        quotes["straight"] += text.count('"')
        not_but += len(NOT_BUT.findall(text))
        chars += len(HANGUL.findall(text))
        clefts += [where] * len(CLEFT.findall(text))
        seen_middot2 = False
        for code, level, rx, hint in rules:
            if code == "MIDDOT_SPACED1" and seen_middot2:
                continue
            m = rx.search(text)
            if not m:
                continue
            if code == "MIDDOT_SPACED2":
                seen_middot2 = True
            s, e = max(0, m.start() - 25), min(len(text), m.end() + 25)
            snippet = ("…" if s else "") + text[s:e].replace("\n", " ") + ("…" if e < len(text) else "")
            hits.append((level, code, where, snippet, hint))
    if len(clefts) >= CLEFT_MIN:
        hits.append(("INFO", "M_CLEFT", "(문서 전체)", f"'핵심은 ~' 분열문 {len(clefts)}회 ({', '.join(clefts[:5])})",
                     "바로 서술하기"))
    if not_but >= NOT_BUT_MIN and not_but * NOT_BUT_PER_CHARS >= chars:
        hits.append(("INFO", "M_NOTBUT", "(문서 전체)", f"'A가 아니라 B' 대구 {not_but}회 / 한글 {chars}자",
                     "하나만 남기고 나머지는 평서문으로"))
    if quotes["curly"] and quotes["straight"]:
        hits.append(("INFO", "QUOTE_MIX", "(문서 전체)",
                     f"곡선 {quotes['curly']} / 곧은 {quotes['straight']}", "문서 기존 관례로 통일"))
    if prose:
        # txt는 줄바꿈 하나로 문단을 나누는 글(웹 폼·HWP에서 복사)이 많아 묶지 않는다
        is_md = path.suffix.lower() == ".md"
        skip = front_matter_end(path) if is_md else 0
        hits.extend(prose_totals(split_paragraphs(paras, is_md, skip), True))
    meaning = sum(1 for h in hits if h[1].startswith("M_"))
    if not show_info:
        hits = [h for h in hits if h[0] != "INFO"]
    return hits, meaning


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--info", action="store_true", help="INFO(판단 필요) 항목도 표시")
    ap.add_argument("--prose", action="store_true", help="서술형 글(리포트·에세이·블로그) 규칙도 검사")
    ap.add_argument("--out", help="리포트를 UTF-8 파일로 저장(콘솔 cp949 회피)")
    ap.add_argument("--heads", action="store_true",
                    help="머리 문장(제목)과 그 아래 첫 단락만 뽑아 보여 준다(뜻 점검·독립 검토 입력용)")
    args = ap.parse_args()

    if args.heads:
        lines = []
        for f in args.files:
            path = Path(f)
            fn = HEADS.get(path.suffix.lower())
            if not path.is_file() or fn is None:
                lines.append(f"[SKIP] {f}: 파일 없음 또는 머리 문장 추출 미지원(xlsx 는 미지원)")
                continue
            lines.append(f"== {path.name} 머리 문장")
            for where, kind, text in fn(path):
                lines.append(f"  [{kind}] {where}: {text}")
        report = "\n".join(lines)
        if args.out:
            Path(args.out).write_text(report + "\n", encoding="utf-8")
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        print(report)
        sys.exit(0)

    lines, total_warn, bad_input = [], 0, False
    for f in args.files:
        path = Path(f)
        if path.name.startswith("~$"):
            continue
        if not path.is_file() or path.suffix.lower() not in EXTRACTORS:
            lines.append(f"[SKIP] {f}: 파일 없음 또는 미지원 확장자")
            bad_input = True
            continue
        try:
            hits, meaning = scan(path, args.info, args.prose)
        except (zipfile.BadZipFile, ET.ParseError, KeyError, UnicodeDecodeError) as e:
            lines.append(f"[ERROR] {f}: {type(e).__name__}: {e}")
            bad_input = True
            continue
        warn = sum(1 for h in hits if h[0] == "WARN")
        total_warn += warn
        info = f"INFO {len(hits) - warn}" if args.info else "INFO 숨김(--info)"
        lines.append(f"== {path.name}: WARN {warn}, {info}, 뜻 신호 {meaning}")
        for level, code, where, snippet, hint in hits:
            lines.append(f"  [{level}] {code} @ {where}: {snippet}\n         -> {hint}")
        if meaning >= MEANING_REVIEW_MIN:
            lines.append(f"  * 뜻 신호 {meaning}건(기준 {MEANING_REVIEW_MIN}건 이상): 독립 검토 대상"
                         " - --heads 출력을 쓴 맥락 밖 검토 에이전트에게 넘긴다")
    lines.append(f"RESULT: {'FAIL' if total_warn else 'PASS'} (WARN {total_warn})")

    report = "\n".join(lines)
    if args.out:
        Path(args.out).write_text(report + "\n", encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print(report)
    sys.exit(2 if bad_input and not total_warn else (1 if total_warn else 0))


if __name__ == "__main__":
    main()
