# -*- coding: utf-8 -*-
"""
워드 테마 샘플 생성기.

워드 전용 테마(WORD_THEMES)마다 2종(보고서 / 메모)의 .docx 미리보기를 만든다.
    python build_preview.py
  → preview_<theme>_report.docx, preview_<theme>_memo.docx

제목부 스타일 6종 비교 샘플은 title_samples.py 참조.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from word_theme import (new_doc, add_title, add_heading, add_paragraph,
                        add_bullets, add_callout, add_table)
from word_themes import WORD_THEMES

FIN_HEADERS = ["계정과목", "전기", "당기", "증감액", "증감율"]
FIN_ROWS = [
    ["매출액", 1_200_000, 1_450_000, 250_000, 0.2083],
    ["매출원가", -800_000, -950_000, -150_000, 0.1875],
    ["매출총이익", 400_000, 500_000, 100_000, 0.25],
    ["판매관리비", -180_000, -210_000, -30_000, 0.1667],
    ["영업이익", 220_000, 290_000, 70_000, 0.3182],
]
NUMCOLS = {1: "accounting", 2: "accounting", 3: "accounting", 4: "percent"}


def build_report(theme, out):
    doc = new_doc(theme, doc_type="report")   # 테마·문서유형은 문서에 기억 → 이후 자동 상속
    add_title(doc, "FY25 손상평가 보고서", f"테마: {theme} · 2026.03",
              meta={"작성": "홍길동", "검토": "김철수", "일자": "2026-03-15"})
    add_heading(doc, "1. 평가 개요", level=1)
    add_paragraph(doc, "본 보고서는 현금창출단위(CGU)별 손상징후를 검토하고, "
                       "회수가능액과 장부금액을 비교하여 손상차손 인식 여부를 판단한 결과를 요약한다.")
    add_bullets(doc, [
        "대상: 영업권 및 유형자산이 배분된 4개 CGU",
        "방법: 사용가치(DCF)와 처분부대원가 차감 공정가치 중 큰 금액",
        ("핵심 가정: 할인율(WACC), 영구성장률, 추정 현금흐름"),
        (1, "할인율은 외부 시장데이터 기반 산정"),
        (1, "추정기간은 5개년 + 영구"),
    ])
    add_callout(doc, "핵심 결론: 일부 CGU에서 회수가능액이 장부금액을 하회하여 손상차손을 인식하였다.")
    add_heading(doc, "2. 주요 재무지표", level=1)
    add_heading(doc, "2.1 손익 요약", level=2)
    add_paragraph(doc, "전기 대비 당기 손익 변동은 아래와 같다.")
    add_table(doc, FIN_HEADERS, FIN_ROWS, number_cols=NUMCOLS, total_last_row=True)
    add_heading(doc, "3. 평가 절차", level=1)
    add_bullets(doc, [
        "손상징후 식별 및 CGU 식별",
        "회수가능액 산정 (사용가치·공정가치)",
        "장부금액과 비교하여 손상차손 인식",
        "민감도 분석 및 결론 도출",
    ], numbered=True)
    add_heading(doc, "4. 가정 및 한계", level=1)
    add_paragraph(doc, "할인율은 WACC를 적용하였으며, 추정에는 불확실성이 존재한다.")
    doc.save(out)
    print("  생성:", out)


def build_memo(theme, out):
    doc = new_doc(theme, doc_type="memo")
    add_title(doc, "내부 메모: 분기 마감 일정", f"테마: {theme}")
    add_paragraph(doc, "수신: 재무팀 / 발신: 결산담당 / 일자: 2026-06-04")
    add_heading(doc, "안건", level=2)
    add_paragraph(doc, "FY26 1분기 결산 마감 일정을 아래와 같이 공유합니다.")
    add_table(doc, ["단계", "담당", "기한"],
              [["전표 마감", "각 팀", "06-10"],
               ["계정 조정", "결산팀", "06-12"],
               ["검토/승인", "CFO", "06-14"]])
    add_callout(doc, "기한 엄수 요망. 지연 시 사전 공유 바랍니다.")
    doc.save(out)
    print("  생성:", out)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    print("워드 샘플 생성 중...")
    for name in WORD_THEMES:
        build_report(name, os.path.join(here, f"preview_{name}_report.docx"))
        build_memo(name, os.path.join(here, f"preview_{name}_memo.docx"))
    print("완료.")


if __name__ == "__main__":
    main()
