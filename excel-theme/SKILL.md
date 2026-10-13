---
name: excel-theme
description: 엑셀(.xlsx) 파일을 새로 생성하거나 편집할 때 일관된 색상·폰트·셀 서식 테마를 적용한다. openpyxl로 스프레드시트를 만들거나 표를 스타일링할 때, 또는 사용자가 "엑셀 만들어줘 / 표 정리해줘 / 양식 통일" 등을 요청할 때 사용한다.
---

# Excel Theme

Claude Code로 엑셀(.xlsx)을 **생성하거나 편집할 때마다** 이 스킬의 헬퍼를 거쳐
일관된 색상·폰트·셀 서식을 적용한다. 매번 스타일을 새로 짜지 말고 헬퍼를 호출한다.

색/폰트는 이 파일 `excel_theme.py` 상단 **편집 영역(PALETTES)** 이 단일 출처다.
공식 테마는 4종: **`default`**(기본, 정산표·결산 회색 삼선표) /
**`audit`**(감사조서 — 색·서식은 default 와 완전히 같고, 시트 제목 밴드 대신 조서 헤더 6항목을 두며 섹션 소제목을 메인헤더 밴드 `#393939`로 쓴다) /
**`procpa`**(브랜드 블루 — 딥네이비 `#0B1C4A` + 블루 `#2563EB`, procpa.co.kr 팔레트) /
**`dcf-valuation`**(DCF·평가, Ocean Blue 딥블루 `#004889`).
**감사조서(계정조서 4000~6500·중간감사·기말감사 조서)는 `theme="audit"`** — 별칭 `감사조서`/`workpaper`.
`01. Business` 감사 작업공간은 환경변수 `EXCEL_THEME_DEFAULT=audit` 로 기본 테마를 audit 으로 둔다.
(`navy`·`charcoal` 은 legacy — 기존 감사조서 하위호환용으로만 유지, 신규 사용 금지.
구 이름 `closing`·`valuation`·`audit_charcoal`·`audit_navy` 는 별칭으로 계속 동작)

## 언제 적용하나
- openpyxl 등으로 새 .xlsx 를 만들 때 → 표를 채운 뒤 마지막에 `apply_theme(ws, ...)` 호출.
- 기존 .xlsx 의 표 서식을 다듬을 때도 동일하게 `apply_theme` 적용.
- 미지정 시 **기본 `default`**(정산표·결산 회색; `EXCEL_THEME_DEFAULT` 환경변수로 덮어쓰기 가능).
  **감사조서는 `audit`**, 브랜드 산출물은 `procpa`, DCF·평가는 `dcf-valuation` 명시.

## 사용 방법

이 스킬 폴더(`excel-theme/`)를 import 경로에 추가하고 `apply_theme` 를 호출한다.

```python
import sys
sys.path.insert(0, r"C:\Users\PC\.claude\skills\excel-theme")

from openpyxl import Workbook
from excel_theme import apply_theme, style_total_row, THEMES

wb = Workbook()
ws = wb.active
# ... 헤더와 데이터를 먼저 채운다 ...

apply_theme(
    ws,
    theme="default",            # THEMES 키 중 하나(default/procpa/dcf-valuation). 생략 시 기본값.
    header_row=4,               # 헤더가 있는 행 번호
    data_range="B4:F9",         # 표 영역. 생략하면 사용 영역 자동 감지.
    title_cell="B2",            # (선택) 시트 제목 셀 — 밴드 서식. 내용은 **시트명만** (아래 기본규칙 8)
    number_format_cols={        # (선택) 열별 숫자 서식 (회계서식 권장)
        "C": "accounting",      #   숫자: 천단위 콤마, 음수 빨강 괄호, 0→대시
        "F": "percent_acct",    #   퍼센트: 음수 빨강 괄호, 0→대시
    },
)
# 소계/합계 행은 별도 강조 (medium 상단선 + 굵게)
style_total_row(ws, THEMES["default"], row=9, min_col=2, max_col=6)

wb.save("output.xlsx")
```

`apply_theme` 는 아래 4종 셀 규칙을 한 번에 적용한다. 개별 제어가 필요하면
인자 `zebra=`, `freeze=`, `autofit=`, `highlight_neg=` 로 끄거나, 개별 함수를 직접 호출한다.

### 더 간단히: `write_table` (권장 — 데이터+옵션 한 번에)
셀을 일일이 채우는 대신, 데이터와 옵션만 주면 채우기·서식·소계/합계·입력색까지 끝난다:
```python
from excel_theme import write_table
write_table(ws, ["계정","당기","전기"], data, theme="default",
            title="재무상태표",
            currency_cols=[1, 2],                 # 통화 열 → 원 단위 회계서식(accounting)
            section_rows=[1], subtotal_rows=[4], total_rows=[5],   # rows의 1-based 인덱스
            input_cells=["C6:D7"], linked_cells=["C10:D10"])       # 입력=크림 / 연결=파랑
```
`mark_cells(ws, theme, input=…, linked=…)` 로 입력(#FFFBEF)·연결(#CCECFF) 셀만 따로 칠할 수도 있다.

## 사용 가능한 테마 (공식 4종 + legacy 2종)
| 이름 | 분위기 | 본문 헤더 | 표 스타일 |
|---|---|---|---|
| `default` (기본) | 정산표·결산 — 회색 | #F2F2F2/검정 | frame(삼선표) |
| `audit` | 감사조서 — default 색 + 소제목=메인헤더 밴드(#393939) | #F2F2F2/검정 (서브표 #DDE3E8) | frame(삼선표, default 동일) |
| `procpa` | 브랜드 블루 — 딥네이비 & 블루 | #EFF5FF/딥네이비 | frame(삼선표) |
| `dcf-valuation` | DCF·평가 — Ocean Blue | #004889/흰색 | rules |
| `navy` *(legacy)* | 네이비 & 골드 | #E1E8F1/네이비 | frame(삼선표) |
| `charcoal` *(legacy)* | 차콜 & 틸 | #D6DEE6/차콜 | frame(삼선표) |

> **구 이름 `valuation`→`dcf-valuation`, `closing`→`default`, `audit_charcoal`→`charcoal`, `audit_navy`→`navy` 는 별칭(ALIASES)으로 계속 동작** — 감사조서 빌더 `frame.py`·valuation-tools 등 하위호환. legacy 2종은 신규 산출물에 쓰지 않는다(미리보기 빌드에서도 제외).
> **default·procpa(+legacy navy·charcoal) 공통 규칙(dcf-valuation 제외):**
> - 표 = **frame(삼선표)**: 좌/우 외곽선 없음 · 상/하 **굵게(medium)** · 내부 **얇게(thin)**.
> - 셀 배경: 하드코딩(입력)=`note`(#FFFBEF) · 수식=흰배경 · 참조값=`linked`(#F2F2F2).
> - 숫자=`accounting`(`#,###,##0;[Red](#,###,##0);-`) · 비율=`percent_acct`(`0.00%_);[Red](0.00%);"-"_);@_)`).
> - 헤더 3색: **제목**(B2, `title_bg`) / **본문 기본**(`excel_header_fill`) / **본문 세컨더리**(`excel_header2_fill`, 2차 표용).
>   세컨더리 헤더는 `style_header_row(..., secondary=True)` 로 적용.
> - **섹션행 배경색 없음**(`subheader=#FFFFFF`, 굵게만). **소계·합계 배경**: `default`·`audit`=**헤더색 #F2F2F2**, `procpa`=**#EFF5FF**(`excel_total_fill`), legacy `charcoal`·`navy`=없음(윗선+굵게만). dcf-valuation=band 유지.
> - **음수**: 테마 `excel_highlight_neg=False` → 굵게/별색 오버레이 없이 **숫자서식 `[Red]`(기본 빨강)만**. `apply_theme`/`write_table` 가 테마값을 따르므로 실제 사용에서도 동일.

색상·폰트·서식·디자인 원칙 상세는 `GUIDE.md` 참고.
**테마 색을 바꾸려면 `excel_theme.py` 상단 PALETTES 만 수정**하면 엑셀·워드·PPT가 함께 바뀐다.

## 셀 규칙 → 함수 매핑 (`excel_theme.py`)
1. **숫자/통화 서식** — `set_number_formats` : 회계서식 `accounting`(#,###,##0;[Red](#,###,##0);-, **0은 대시**), `percent_acct`, 원화/백만원, 소수 등.
2. **테두리·정렬** — `apply_borders`(외곽 굵게+내부 얇게), `style_body`(숫자 우측·텍스트 좌측, 세로 가운데). 행높이는 건드리지 않는다.
3. **헤더·틀고정·줄무늬** — `style_header_row`(배경색+글씨), `freeze_below_header`, `zebra_stripes`.
4. **자동 열너비·조건부서식** — `autofit_columns`(한글 폭 보정), `highlight_negatives`/`highlight_threshold`.
5. **정산표 관행** — `style_total_row`(소계/합계 행: medium 상단선+thin 하단선+굵게+배경), `style_subheader_row`(섹션 구분 행).
6. **시트 탭 색** — `set_tab(ws, role, theme)` : 역할(`guide`/`output`/`calc`/`input`/`pbc`/`raw`) → 테마 `tab_colors`. `write_title`·`write_audit_header` 의 `tab=` 인자로도 지정.

## 계산값은 반드시 수식으로 (사용자 확정 2026-09-04 · 필수)
**합계·소계·차이·비율·건수·표본수 등 다른 셀에서 계산되는 값은 절대 파이썬에서 계산한 숫자를 값으로 쓰지 않고
엑셀 수식(SUM/SUMIFS/COUNTIFS/ROUND/IF …)으로 넣는다.** 집계의 원천 데이터가 조서 밖에 있으면 원천 라인을
`Data_*` 시트로 함께 넣고 그 시트를 참조한다. 값으로 들어가는 것은 ① PBC·외부자료 원자료(입력=크림 `note_fill`)
② 판단·가정 파라미터(PM·RF·환산계수 등, 입력=크림) ③ 텍스트·문서번호뿐이다.
이유: 조서는 검토자가 셀을 눌러 계산 근거를 추적할 수 있어야 하고, 입력값이 바뀌면 자동 재계산돼야 한다.
산출 직후 `theme_lint`와 별도로 "숫자인데 수식이 아닌 셀"을 훑어 계산값이 섞이지 않았는지 확인한다.

## 엑셀 생성 기본 규칙 (apply_theme 가 자동 적용)
1. **틀 고정 금지** — `freeze=True`·`ws.freeze_panes` 설정을 하지 않는다(사용자 확정 2026-07).
   기존 파일을 편집할 때도 틀고정을 추가하지 말고, 발견하면 제거한다. theme_lint가 위반을 잡는다.
2. **A열·1행 여백** — `margin=True`(기본). A열 너비 2.0 으로 비워 둠(1행은 비워 두되 높이는 기본값). 표는 B2 제목/B열부터.
3. **제목 색 확장** — `title_cell="B2"` 의 색(채움)을 **데이터가 있는 마지막 열까지** 가로로 확장.
   제목 셀에 들어가는 **내용은 시트명뿐**이다(기본 규칙 8).
4. **표 소제목(캡션)** — `write_caption(ws, "B12", "표1. …")` 로만 쓴다: **맑은 고딕 11pt 굵게·검정**(회색·10pt 금지).
   캡션 행 **바로 아래 1행을 비우고** 그 다음 행에 표 헤더(`header_row = 캡션행 + 2`). 사용자 확정 2026-09-04.
   **캡션은 짧은 제목만 (사용자 확정 2026-09-12 · 필수)** — 표가 무엇인지 알리는 명사구 한 줄로 끝낸다(`MAX_CAPTION` 40자).
   계산 방법·근거·판단 사유·주의사항을 캡션에 이어 붙이지 않는다
   ("2-1 재질별 기준 단가 - 재료 단가는 전기 결산 단가나 회사 단가표로, 가공 단가는 전기 가공비 ÷ 생산 kg로 정한다…" ✗
    → "2-1 재질별 기준 단가" ✓). 설명이 필요하면 **표의 비고·근거 열**이나 별도 안내 시트에 쓴다.
   시트 상단에 설명 문단을 여러 줄 까는 것도 같은 이유로 지양한다. theme_lint 가 "캡션이 너무 김"으로 잡는다.
4. **정렬 고정** — 숫자·비율·수식 셀은 **오른쪽 정렬**, 텍스트는 왼쪽(사용자 확정 2026-09-04). `style_body`/`set_number_formats`가 자동 적용하며, 문자열을 반환하는 수식(IF(…,"OK","확인")·TEXT·&)만 텍스트로 취급한다. SUMIFS/COUNTIFS의 조건 문자열("자영")이나 IF(…,"",…)의 빈 문자열은 텍스트 반환이 아니므로 오른쪽이며, 숫자 서식(`number_cols`)을 지정한 열의 수식은 무조건 오른쪽이다(2026-09-04 보강). 본문 폰트는 **맑은 고딕 11pt 검정** 고정.
5. **열너비 상한 `MAX_COL_WIDTH`=30** (사용자 확정 2026-09-08) — `autofit_columns` 는 이보다 넓히지 않는다.
   한 열을 길게 늘리지 말고, 긴 문장은 **wrap_text** 또는 표 폭까지 **병합**으로 처리한다.
   조서 헤더·캡션·설명문 등 표 밖 텍스트는 열너비 산정에 넣지 않는다.
6. **행높이는 건드리지 않는다 (사용자 확정 2026-09-09 · 필수)** — 어떤 행에도 `row_dimensions[r].height` 를 지정하지 않고
   **엑셀 기본 행높이**를 그대로 둔다. 헤더·밴드·조서 헤더·1행 여백·줄바꿈 셀 모두 예외 없음. `fit_rows` 는 폐기(호출해도 아무 것도 하지 않음).
   긴 문장은 wrap_text 만 켜고 열너비·병합으로 처리하며, 줄 수에 맞춘 행높이 계산을 하지 않는다.
   기존 파일을 편집할 때도 행높이를 새로 넣지 않는다. theme_lint 가 "행높이 지정 위반"으로 잡는다.
7. **시트 탭 색 = 시트의 역할 (사용자 확정 2026-09-10 · 필수)** — `set_tab(ws, role, theme)` 또는
   `write_title(..., tab="calc")` / `write_audit_header(..., tab="output")` 로만 칠한다. 직접 `tabColor` 를 넣지 않는다.
   역할 6종: `guide`(안내·표지 — **색 없음**) · `output`(결과·총괄) · `calc`(계산·집계) ·
   `input`(입력 — 작성자) · `pbc`(입력 — 회사 제공) · `raw`(원본·참고). 한글·별칭 허용(`회사`·`input_ext` → pbc 등).
   두 축: **진하기 = 자리**(진할수록 파일의 중심인 결과, 연할수록 바깥인 원천 자료),
   **무채색 = 우리 산출물이 아님**(`raw` 는 어느 테마에서든 #D9D9D9 고정).
   색은 그 테마가 이미 가진 토큰에서만 고른다 — 탭을 위해 새 색을 만들지 않는다.
   회색 테마(default·audit·legacy)는 Office 표준 회색 램프 `393939 / 808080 / A6A6A6 / BFBFBF / D9D9D9`,
   색 있는 테마는 팔레트 `tabs` 로 오버라이드(procpa `0B1C4A / 2563EB / 94A3B8 / C9D4E5`,
   dcf-valuation `004889 / 2B579A / 5B9BD5 / D6E4F0`).
   탭 색은 **성격만** 나타낸다 — 진행 상태(완료·검토중)·임시 표시에 쓰지 않고, 한 파일에서 같은 역할은 같은 색이다.
   시트 순서는 안내 → 결과 → 계산 → 입력 → 원본(회사에 양식으로 보내는 파일만 입력을 앞으로).
   기존 파일을 편집할 때는 탭 색을 새로 칠하지 않고, 이미 칠해져 있으면 규약 색으로 맞춘다.
   theme_lint 가 "탭 색 규약 이탈"로 잡는다. 미리보기: `examples/preview_tab_<테마>.xlsx`.
8. **시트 제목 밴드는 시트명만 (사용자 확정 2026-09-10 · 필수)** — `default` 계열 테마의 시트 상단 밴드(B2)에는
   **시트 이름 하나만** 넣는다. 회사명·기준일·목적·괄호 설명 같은 문구를 제목에 덧붙이지 않는다
   ("20 연령분석 — 매출채권 연령별 잔액 검토" ✗ → "연령분석" ✓).
   `write_title(ws, theme=...)` 를 **title 인자 없이** 부르면 시트명(앞 순번을 뗀 `ws.title`)이 자동으로 들어간다
   (`sheet_label`). 문자열을 주더라도 `title_only` 가 뒤의 설명(— – | : · ( [ / 뒤)을 잘라 낸다.
   `write_header(ws, meta, ...)` 의 제목 밴드도 같다(title 생략 → 시트명).
   설명이 필요하면 제목이 아니라 밴드 아래 `write_caption` 또는 본문 첫 줄에 적는다.
   **감사조서 테마 `audit` 은 해당 없음** — 조서 헤더 6항목을 그대로 쓴다(아래 audit 절).
   theme_lint 가 "제목 밴드는 시트명만"으로 잡는다(audit 테마는 검사 제외).
9. **문장부호: 긴 대시 `—` → `-` (사용자 확정 2026-09-11 · 필수, 모든 테마 — 워드·PPT 테마도 동일)** — 엑셀에 들어가는
   **모든 텍스트(제목·캡션·헤더·본문·조서 헤더 값·시트명)** 에 긴 대시(`—` `–` `―`)를 쓰지 않고 하이픈 `-` 을 쓴다
   (`"연령분석 — 잔액"` ✗ → `"연령분석 - 잔액"` ✓, `"1–3월"` → `"1-3월"`). 가운뎃점 `·` 은 그대로 써도 된다.
   헬퍼(`apply_theme`·`write_table`·`write_title`·`write_header`·`write_audit_header`·`write_caption`·`write_section_bar`)가
   `plain_punct` 로 자동 치환하지만, 파이썬에서 문자열을 만들 때부터 `-` 로 쓴다.
   헬퍼 밖에서 셀에 직접 쓰는 텍스트는 `plain_punct(text)` 를 거친다. **수식(`=…`)은 자동 치환하지 않는다**
   (시트 참조가 깨질 수 있음) — 수식 안 문자열도 처음부터 `-` 로 쓴다. 시트명도 자동으로 바꾸지 않으므로 처음부터 `-` 로 짓는다.
   기존 파일을 편집할 때는 수정 범위 안의 텍스트만 맞춘다. theme_lint 가 "문장부호 위반"으로 잡는다.


## 폰트 (앱별 변형)
엑셀 기본 폰트는 **맑은 고딕**이다(procpa·dcf-valuation 의 워드/PPT는 Pretendard).
엑셀 폰트를 바꾸려면 `excel_theme.py` 의 해당 테마 `font_excel` 값을 수정한다.

## 표 유형별 샘플 + 규칙 커스터마이징
`examples/table_samples.py` 의 `SAMPLES` 리스트에 표 유형이 선언돼 있다(단순표 / 합계만 /
소계+합계 / 섹션+소계+합계 / 묶음헤더 / 와이드+퍼센트 / 줄무늬OFF / 2열목록 / 입력·연결셀 등).
각 표의 규칙(`subheader_rows`, `subtotal_rows`, `total_rows`, `number_format_cols`,
`zebra`, `freeze` 등)을 dict 값만 고쳐 직접 커스터마이징한다.

## 디자인 원칙
공통 디자인 규칙(절제된 악센트·헤어라인·60-30-10·anti-slop)은 `GUIDE.md`(C. 디자인 원칙) 를 따른다.

## 검증
`python examples/build_preview.py` → 테마마다 `preview_<theme>.xlsx`(샘플마다 한 시트) 생성.
전체(엑셀·워드·PPT) 일괄 미리보기는 `build_all.py`(이 폴더).

**산출 직후 필수 체크**: `python theme_lint.py <산출.xlsx> [테마] [시트…]` → **위반 0 확인 후 납품**.
(테마 외 폰트·하드코딩 배경색·규약 이탈을 잡는다.) 시트를 손스타일로 만들지 말 것 —
반드시 `apply_theme`/`write_table`/`style_*` 헬퍼 경유. 제목은 `title_cell=` 규약(`title_bg` #393939 밴드+흰 **11pt** — 조서 헤더·제목 밴드 11pt, 사용자 확정 2026-09-07),
강조 배경은 `note_fill`(#FFFBEF)·`linked_fill`(#F2F2F2) 토큰만 사용(초록·주황 등 임의 색 금지), 본문 폰트 11pt.
탭 색은 `set_tab` 역할 색만(임의 색 금지, 안내·표지는 색 없음).

> 참고: `tools/` 폴더는 테마 *개발용* 도구(샘플 피드백 diff·실무파일 분석기)다. 엑셀 생성 런타임에서는 사용하지 않는다 — 개발 워크플로우는 프로젝트 폴더의 DEV.md 참조.

## 감사조서 테마 `audit` (사용자 확정 2026-09-08)
**색·폰트·표 규칙은 default 와 완전히 같다**(틱마크 빨강 `accent`만 다름). 감사조서용으로 다른 것은 아래 "쓰는 법" 두 가지다.
default 와 다른 점은 두 가지뿐이며, 감사조서를 만들 때 **반드시** 지킨다:
1. **시트 제목 밴드(B2)는 두지 않는다.** 시트 제목은 조서 헤더의 '조서명'에 적는다. **섹션 소제목은 메인헤더 밴드** —
   `write_section_bar(ws, row, "1. 총괄표", theme="audit", last_col=8)` → `#393939`+흰 굵은 글씨.
   표준조서의 "1. 목적 / 2. 총괄표 / 3. Test 수행" 줄이 이것. 반환값(row+2)이 다음 표 `header_row`.
2. **시트마다 상단에 조서 헤더 6항목** — `hdr = write_audit_header(ws, meta, theme="audit")` → 반환 5(4행 공백).
   B2:G3 에 **[회사명|조서번호|시트명] / [결산일|조서명|작성자]**, 병합 없음, 색 없음(라벨 굵게만),
   위아래 굵은 선·가운데 얇은 선·좌우 외곽선 없음. `meta = {"회사명","조서번호"(시트번호 포함, 예 "6000B-10"),
   "결산일","조서명"(조서·파일 이름),"작성자"}`, 시트명은 기본 `ws.title` 에서 **앞 순번을 뺀 이름**("10 개요"→"개요", 확정 2026-09-09).
   **표지 시트는 두지 않는다**(`write_cover_sheet` 는 legacy, 신규 사용 금지).
   **헤더 값 때문에 열너비를 늘리지 않는다.**
3. **시트 이름 = 조서번호 순번 + 이름** — 조서 헤더가 있는 시트는 `10 총괄표`, `20 세부Test` 처럼 시트명 앞에 10 단위 번호를
   붙이고 헤더 조서번호는 `6000A-10` 으로 맞춘다. Sampling 관련 시트(MUS Sample·모집단·제외전표·원장 raw)는 번호를 붙이지 않고
   조서번호도 파일 번호(`6000A`)만 쓴다.
- 표: `write_table(..., theme="audit", header_row=hdr)` → default 와 같은 회색 헤더·합계. 서브표는
  `write_table(..., secondary=True)` → 헤더·합계 모두 `#DDE3E8`. 표 맨 아래 합계행은 하단 medium 선 유지.
- 틱마크 `write_tickmark(ws, "C12", "R", theme="audit", note="전기조서 대사")` → 빨강 글씨·가운데
  (R 전기조서 · GL 총계정원장 · F Footing · PR 제시재무제표 · Ref. 조서번호).
- 통화 `currency_cols` 는 원 단위 회계서식(`accounting`, 사용자 확정 2026-09-14). 백만원 표시가 필요하면 `number_cols={…:"million_won"}` 로 지정.
- 린트: `python theme_lint.py <파일> audit`.

