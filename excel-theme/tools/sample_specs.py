# -*- coding: utf-8 -*-
"""
■ 피드백용 추가 샘플 명세 — 기존 examples/table_samples.py 의 SAMPLES 8종에
  더해, 피드백 루프에서만 쓰는 4종을 정의한다 ■

필드는 table_samples.py 와 동일 + 추가 필드:
  input_cells   : 직접 입력(하드코딩) 셀/범위 → 크림(#FFFBEF). 시트 좌표 기준.
  linked_cells  : 타시트 연결 셀/범위 → 파랑(#CCECFF).
  tables        : [spec, spec] — 한 시트에 표 여러 개 (12_복수표 전용).
                  두 번째 표부터는 제목 없이 첫 표 아래 2행 간격으로 배치.
"""

ACC = "accounting"
PCT = "percent_acct"
MUL = "multiple"
MWN = "million_won"

EXTRA_SAMPLES = [
    # ── 9. 입력/연결 셀 마킹 (DCF·조서 관행) ─────────────────────
    {
        "name": "09_입력연결셀",
        "title": "주요 가정 (입력=크림 / 연결=파랑)",
        "headers": ["항목", "값", "비고"],
        "rows": [
            ["매출성장률", 0.05, "경영계획"],
            ["WACC", 0.085, "WACC 시트"],
            ["영구성장률", 0.01, "거시지표"],
            ["법인세율", 0.209, "세법"],
        ],
        "number_format_cols": {"C": PCT},
        # 데이터는 B5~D8. 값(C열) 중 직접입력 3개=크림, 타시트 연결 1개=파랑.
        "input_cells": ["C5", "C7", "C8"],
        "linked_cells": ["C6"],
    },

    # ── 10. 평가 요약 (배수·퍼센트 서식, valuation 실전형) ────────
    {
        "name": "10_평가요약",
        "title": "비교회사 멀티플 요약",
        "headers": ["비교회사", "EV/EBITDA", "PER", "PBR", "영업이익률"],
        "rows": [
            ["A사", 8.5, 12.3, 1.4, 0.082],
            ["B사", 7.2, 10.8, 1.1, 0.065],
            ["C사", 9.1, 14.2, 1.8, 0.094],
            ["평균", 8.3, 12.4, 1.4, 0.080],
            ["중앙값", 8.5, 12.3, 1.4, 0.082],
        ],
        "number_format_cols": {"C": MUL, "D": MUL, "E": MUL, "F": PCT},
        "subtotal_rows": [4],
        "total_rows": [5],
    },

    # ── 11. 정산표 미니 (statement 룩 + 묶음헤더 + 백만원) ────────
    {
        "name": "11_정산표미니",
        "title": "재무상태표 (요약)",
        "group_header": [("계정", 1), ("금액(백만원)", 2)],
        "headers": ["계정과목", "당기", "전기"],
        "rows": [
            ["자산", None, None],
            ["유동자산", 45_200_000_000, 41_800_000_000],
            ["비유동자산", 88_400_000_000, 86_100_000_000],
            ["자산총계", 133_600_000_000, 127_900_000_000],
        ],
        "number_format_cols": {"C": MWN, "D": MWN},
        "subheader_rows": [1],
        "total_rows": [4],
        # 전기 수치는 확정값 직접입력 관행
        "input_cells": ["D7:D8"],
    },

    # ── 12. 한 시트에 표 2개 (표 간 간격·제목 규칙 검증) ──────────
    {
        "name": "12_복수표",
        "title": "표 2개 배치",
        "tables": [
            {
                "headers": ["구분", "금액"],
                "rows": [
                    ["수익", 1_450_000],
                    ["비용", -1_160_000],
                    ["이익", 290_000],
                ],
                "number_format_cols": {"C": ACC},
                "total_rows": [3],
            },
            {
                "headers": ["항목", "당기", "전기"],
                "rows": [
                    ["인원수", 142, 138],
                    ["평균급여", 68_000_000, 65_000_000],
                ],
                "number_format_cols": {"C": "thousands", "D": "thousands"},
            },
        ],
    },
]
