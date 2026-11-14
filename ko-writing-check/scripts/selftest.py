# -*- coding: utf-8 -*-
"""slop_scan.py 자체 시험.

구호·번역투 문장은 걸리고, 평범한 업무 문장은 걸리지 않는지 확인한다.
스캐너 규칙을 바꾼 뒤 돌린다.

usage: python selftest.py
exit : 0 = 모두 통과, 1 = 실패 있음
"""
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import slop_scan  # noqa: E402

# (문장, 반드시 나와야 하는 코드)
MUST_HIT = [
    ("전표에서 공시까지 만들고, 공시를 읽고, 그 과정을 지킵니다", {"M_FROMTO", "M_TRIPLE"}),
    ("AI가 작성하는 시대, 재무담당자는 처음부터 검토자", {"M_ERA"}),
    ("사원과 과장이 수준에 맞게 배웁니다", {"M_PROMISE"}),
    ("작성자에서 검토자로", {"M_TRANSFORM"}),
    ("실무 역량을 강화하고 활용도를 제고합니다", {"M_POLICY"}),
    ("이번 개정은 리스 실무의 새로운 지평을 열 것입니다", {"M_CLICHE"}),
    ("전문가들은 리스 판단이 어렵다고 말합니다", {"M_HEARSAY"}),
    ("좋은 질문입니다. 아래에 정리해 드리겠습니다", {"M_CHATBOT"}),
    ("결산 일정을 앞당긴 것은 결국 감사 대응 때문이라는 이유다", {"M_CONCLUDE"}),
    ("매출 — 전기 대비 증가", {"EMDASH"}),
    ("매출 · 매입 · 재고를 대사함", {"MIDDOT_SPACED2"}),
    ("결산 업무에 있어서 중요한 절차임", {"TL_ISSEO"}),
    ("회수가능성은 낮다고 판단되어진다", {"TL_DOEEOJI"}),
    ("40,000 은 미수금으로 계상함", {"PARTICLE_SPACE"}),
    ("연체 채권이 늘어난 것으로 보여집니다", {"TL_BOYEOJI"}),
]

# 평범한 업무 문장: WARN 0, 뜻 신호 0 이어야 한다
MUST_PASS = [
    "매출채권은 전기말 대비 12% 증가하였으며, 주요 원인은 A사 외상매출 증가임",
    "수강생은 AI가 써 둔 분개 20건을 채점하고 틀린 5건을 고칩니다",
    "1일차에는 분개부터 현금흐름표까지 만들고, 2일차에는 상장사 공시를 분석합니다",
    "내부회계관리제도 구축 용역을 2025년에 수행함",
    "대손충당금에 대한 회계정책은 전기와 동일하게 적용함",
    "요청드린 자료 중 매출·매입 원장은 수령하였습니다",
    "첨부 파일의 3번 시트에서 수식을 확인해 주시기 바랍니다",
    "리스부채는 리스개시일에 지급되지 않은 리스료의 현재가치로 측정함",
    "회의는 11월 18일 오후 2시에 3층 회의실에서 진행합니다",
    "그렇다면 이 차이는 어디에서 생겼을까요? 원인은 기간 귀속입니다",
]

# 문서 단위 기준: 한 번은 안 걸리고, 여러 번이면 걸린다
CLEFT_ONE = "핵심은 증빙과 대조하는 절차입니다"
CLEFT_MANY = "핵심은 검토입니다\n핵심은 증빙입니다\n핵심은 대조입니다"


def run(text):
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "case.txt"
        p.write_text(text + "\n", encoding="utf-8")
        hits, meaning = slop_scan.scan(p, True)
    return {h[1] for h in hits}, sum(1 for h in hits if h[0] == "WARN"), meaning


def main():
    fails = []
    for text, want in MUST_HIT:
        codes, _warn, _m = run(text)
        if not want <= codes:
            fails.append(f"안 걸림: {text}\n    기대 {sorted(want)} / 실제 {sorted(codes)}")
    for text in MUST_PASS:
        codes, warn, meaning = run(text)
        if warn or meaning:
            fails.append(f"잘못 걸림: {text}\n    실제 {sorted(codes)}")
    codes, _w, meaning = run(CLEFT_ONE)
    if meaning:
        fails.append(f"잘못 걸림(한 번 쓴 '핵심은'): {sorted(codes)}")
    codes, _w, _m = run(CLEFT_MANY)
    if "M_CLEFT" not in codes:
        fails.append("안 걸림: '핵심은' 3회")
    slogans = "\n".join(t for t, want in MUST_HIT[:3])
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "slogans.txt"
        p.write_text(slogans + "\n", encoding="utf-8")
        _hits, meaning = slop_scan.scan(p, False)
    if meaning < slop_scan.MEANING_REVIEW_MIN:
        fails.append(f"구호 세 문장의 뜻 신호가 {meaning}건: 독립 검토 기준({slop_scan.MEANING_REVIEW_MIN}) 미달")

    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    total = len(MUST_HIT) + len(MUST_PASS) + 3
    if fails:
        print("\n".join(fails))
        print(f"FAIL: {len(fails)} / {total}")
        sys.exit(1)
    print(f"PASS: {total}건")


if __name__ == "__main__":
    main()
