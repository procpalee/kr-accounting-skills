# 클로드 앱 · 클로드 코드 - openpyxl 로 .xlsx 만들기

파이썬으로 엑셀 파일을 만들거나 고칠 때마다 `scripts/excel_theme.py` 헬퍼를 거쳐 서식을 입힌다.
매번 스타일을 새로 짜지 말고 헬퍼를 호출한다(클로드 엑셀의 Office.js 경로와 같은 결과가 나온다).

## 1. 불러오기

```python
import sys
sys.path.insert(0, "/mnt/skills/user/excel-theme/scripts")   # 클로드 코드: ~/.claude/skills/excel-theme/scripts
from openpyxl import Workbook
from excel_theme import (write_table, write_title, write_caption, write_section_bar, set_tab,
                         mark_cells, apply_theme, style_total_row, get_theme)
```

## 2. 표 만들기 - `write_table` (권장)

데이터와 옵션만 주면 채우기·서식·소계/합계·입력색까지 한 번에 끝난다.

```python
wb = Workbook(); ws = wb.active; ws.title = "재무상태표"
rows = [
    ["유동자산", 1200, 1100],
    ["비유동자산", 800, 750],
    ["자산총계", "=SUM(C5:C6)", "=SUM(D5:D6)"],     # 합계는 수식으로
]
write_table(ws, ["계정", "당기", "전기"], rows,
            title="재무상태표",            # B2 제목 밴드 - 시트명만
            currency_cols=[1, 2],          # 통화 열(0-based, 항목명 열=0) → 원 단위 회계서식
            total_rows=[3])                # rows 기준 1-based
set_tab(ws, "output")
wb.save("재무상태표.xlsx")
```

| 인자 | 뜻 |
|---|---|
| `title` | B2 제목 밴드 내용(시트명만). `None` 이면 제목 없음 |
| `header_row` / `start_col` | 헤더 행(기본 4) / 시작 열(기본 2 = B) |
| `currency_cols` / `percent_cols` | 통화·비율 열(0-based 데이터열 번호) |
| `number_cols` | 열별 서식 직접 지정 `{2: "million_won"}` 또는 `{"D": "percent_acct"}` |
| `section_rows` / `subtotal_rows` / `total_rows` | rows 기준 1-based 행 번호 |
| `input_cells` / `linked_cells` / `todo_cells` | 입력(크림) · 참조(연회색) · 미입수(노랑) 범위 |
| `secondary` | 한 시트의 두 번째 표 - 헤더·합계를 서브표 색으로 |

숫자서식 키: `accounting`(기본 통화, 음수 빨강 괄호·0은 대시), `million_won`, `million`, `billion`, `thousands`,
`currency_won`, `percent_acct`, `percent`, `decimal`, `multiple`, `change`, `bp`, `date`.

## 3. 시트 요소

- **제목 밴드만**: `write_title(ws, last_col=7, tab="calc")` - title 을 생략하면 시트명(앞 순번 제외)이 들어간다. 반환값 4 = 첫 표 헤더 행.
- **캡션(표 소제목)**: `write_caption(ws, "B12", "표1. 연령별 잔액")` → 헤더는 캡션 행 + 2(한 줄 비움).
- **섹션 밴드**: `write_section_bar(ws, 6, "1. 총괄표", last_col=8)` → 반환값(row+2)이 다음 표 `header_row`.
- **탭 색**: `set_tab(ws, 역할)` - `guide`(안내·표지, 색 없음) · `output`(결과) · `calc`(계산) · `input`(작성자 입력) · `pbc`(회사 제공 자료) · `raw`(원본). 시트 순서는 안내 → 결과 → 계산 → 입력 → 원본.
- **입력·참조 셀만 따로**: `mark_cells(ws, get_theme(), input=["C6:D7"], linked=["C10"])`.
- 이미 채워 둔 표에 서식만 입힐 때: `apply_theme(ws, header_row=4, data_range="B4:F9", title_cell="B2", number_format_cols={"C": "accounting"})`,
  합계행은 `style_total_row(ws, get_theme(), row=9, min_col=2, max_col=6)`.

## 4. 헬퍼가 자동으로 지키는 것
틀고정 없음 · A열 여백(너비 2) · 제목 밴드를 표 끝 열까지 확장 · 숫자·수식 오른쪽 / 텍스트 왼쪽 정렬 ·
열너비 자동(최대 30자) · 행높이 지정 안 함 · 긴 대시 → `-`(수식 제외) · 본문 맑은 고딕 11pt 검정.

## 5. 검증 (산출 직후 필수)
```bash
python /mnt/skills/user/excel-theme/scripts/theme_lint.py 결과.xlsx default
```
위반 0 을 확인한 뒤 넘긴다. 그리고 "숫자인데 수식이 아닌 셀"을 훑어 계산값이 값으로 들어가지 않았는지 본다.
