# -*- coding: utf-8 -*-
"""
write_table / mark_cells 데모.

데이터 + 옵션만 주면 한 번에 themed 표가 완성되는 고수준 API를 보여준다.
    python example_write_table.py
  → 테마별 example_write_table_<theme>.xlsx 생성.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook
from excel_theme import write_table, THEMES

# 현금흐름표 형태 (섹션·소계·합계 + 입력/연결 셀)
DATA = [
    ["영업활동",        None,  None],   # 1 섹션 라벨
    ["당기순이익",       290,   220],   # 2 (입력)
    ["감가상각비",        60,    55],   # 3 (입력)
    ["영업활동 소계",     350,   275],   # 4 소계
    ["투자활동",        None,  None],   # 5 섹션 라벨
    ["유형자산 취득",    -120,   -90],   # 6 (타시트 연결)
    ["투자활동 소계",    -120,   -90],   # 7 소계
    ["현금 순증감",       230,   185],   # 8 합계
]


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    for theme in THEMES:
        wb = Workbook()
        ws = wb.active
        ws.title = theme
        write_table(
            ws, ["항목", "당기", "전기"], DATA, theme=theme,
            title=f"현금흐름표 (테마: {theme})",
            currency_cols=[1, 2],            # 당기·전기 = 테마 통화서식(원 단위 회계서식)
            section_rows=[1, 5],
            subtotal_rows=[4, 7],
            total_rows=[8],
            input_cells=["C6:D7"],           # 당기순이익·감가상각 = 직접 입력(크림)
            linked_cells=["C10:D10"],        # 유형자산 취득 = 타시트 연결(파랑)
        )
        out = os.path.join(here, f"example_write_table_{theme}.xlsx")
        wb.save(out)
        print("생성:", out)


if __name__ == "__main__":
    main()
