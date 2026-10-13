# 엑셀/오피스 테마 가이드 (GUIDE)

색·폰트·표 규칙의 **단일 출처는 `excel_theme.py` 상단 "편집 영역"(`PALETTES`)** 이다.
이 문서는 *사람이 읽는 참고서*(테마 카탈로그 + 엑셀 서식 규격 + 디자인 원칙)일 뿐,
값을 바꾸려면 코드를 수정한다. 호출법·핵심 규칙은 `SKILL.md` 참고.

---

## A. 테마 카탈로그 (공식 4종 + legacy 2종)

| 이름 | 용도 | 표 스타일 | primary | 폰트(워드·PPT / 엑셀) |
|---|---|---|---|---|
| **default** (기본) | 정산표·결산·조서 | statement/frame (회색 헤더·상하 굵은선) | `#404040` | 맑은 고딕 |
| **audit** | 감사조서 | default 와 색·서식 동일, 소제목=메인헤더 밴드 `393939` + 조서 헤더 | `#404040` | 맑은 고딕 11pt |
| **procpa** | 브랜드 산출물 (제안서·보고서) | statement/frame (블루 틴트 헤더) | `#2563EB` / 딥네이비 `#0B1C4A` | Pretendard / 맑은 고딕 |
| **dcf-valuation** | DCF·평가 | rules (흰 배경·가로구분선) | `#004889` Navy | Pretendard / 맑은 고딕 |
| navy *(legacy)* | (구) 감사조서 | frame | `#1F3A5F` | 맑은 고딕 |
| charcoal *(legacy)* | (구) 감사조서 | frame | `#2A2E35` | 맑은 고딕 |

> legacy 2종은 기존 산출물 하위호환용(별칭 `audit_navy`/`audit_charcoal`)으로만 유지 —
> 신규 산출물·미리보기 빌드에서 제외. 구 이름 별칭: `valuation`→`dcf-valuation`, `closing`→`default`.

**테마별 렌더 규칙**
- **audit**: 색·서식은 default 와 동일(틱마크 빨강만 다름). 시트 제목 밴드 없이 상단에 조서 헤더 6항목
  (`write_audit_header`: 회사명·조서번호·시트명 / 결산일·조서명·작성자, B2:G3 병합 없음, 색 없음, 삼선)을 두고,
  섹션 소제목("1. 총괄표" 등)을 메인헤더 밴드(`393939`+흰 글씨, `write_section_bar`)로 쓴다. 첫 표는 5행부터.
  `write_table(secondary=True)`=서브표(헤더·합계 header2 색 통일). 표 맨 아래 합계행은 하단 medium 선을 유지한다.
- **default**: 헤더=회색(`F2F2F2`) bold 검정, 표 **상/하 medium 굵은선**·격자 없음, 타이틀 바(`393939`)가
  데이터 끝열까지 확장, **하드코딩 셀 `FFFBEF`**, 음수=빨강 괄호. (엑셀 헤더색은 워드/PPT와 분리 →
  `excel_header_fill`/`excel_header_text` 오버라이드)
- **procpa**: default 와 같은 statement 레이아웃, 색만 procpa.co.kr 팔레트 — 타이틀 바=딥네이비(`0B1C4A`),
  헤더=블루 틴트(`EFF5FF`)+네이비 글자, 외곽선=네이비, 악센트=블루(`2563EB`).
  입력(`FFFBEF`)·연결(`F2F2F2`)·음수(빨강)는 기능색이라 default 와 동일.
- **dcf-valuation**: 흰 배경(줄무늬 없음), 헤더=흰바탕+네이비 글씨, 음수=괄호+슬레이트(`5A6A7A`, 빨강 아님),
  합계=라이트블루(`F5F8FC`). PPT는 원본 덱 템플릿(`templates/dcf-valuation_source.pptx`)으로 재현.

### 의미색 토큰 (앱마다 다른 곳에 쓰임)
| 토큰 | 쓰임 | default | procpa | dcf-valuation |
|---|---|---|---|---|
| primary | 헤더/제목/타이틀/테두리 | `404040` | `2563EB` | `004889` |
| primary_dark | 주요·소계·합계 글자 | `262626` | `0B1C4A` | `1F3864` |
| secondary / secondary_lt | 보조 | `808080`/`BFBFBF` | `4A5160`/`94A3B8` | `2B579A`/`5B9BD5` |
| band | 합계/줄무늬 배경 | `F2F2F2` | `EFF5FF` | `F5F8FC` |
| subheader | 섹션 구분행 | `FFFFFF`(없음) | `FFFFFF`(없음) | `F5F8FC` |
| note | 콜아웃/하드코딩 셀 | `FFFBEF` | `FFFBEF` | `F5F8FC` |
| accent | 포인트 | `C00000` | `2563EB` | `5B9BD5` |
| negative | 음수 | `FF0000` | `FF0000` | `5A6A7A` |
| text | 본문 글자 | `000000` | `000000` | `3D4F5F` |
| header_font | 헤더 글자(워드/PPT) | `FFFFFF` | `FFFFFF` | `FFFFFF` |

> 엑셀 전용 오버라이드(default·procpa): `excel_header_fill`(헤더 채움)·`excel_header_text`(헤더 글자)·
> `excel_header2_fill`(세컨더리 헤더)·`excel_total_fill`(소계/합계 배경)·
> `title_bg`/`title_font_color`(타이틀 바). 폰트는 `font_excel`(엑셀)·`font_name`(워드/PPT), 크기는 `sizes`.
> 워드 전용 테마 색은 `../word-theme/word_themes.py`(WORD_THEMES)에서 독립 정의 — 여기 팔레트는 폴백용.

---

## B. 엑셀 셀 서식 규격

### B-0. 고수준 표 생성 `write_table` (권장)
데이터+옵션만 주면 채우기·서식·소계/합계·입력색까지 **한 번에**:
```python
write_table(ws, ["계정","당기","전기"], data, theme="default",
            title="재무상태표", currency_cols=[1,2],
            section_rows=[1,5], subtotal_rows=[4,7], total_rows=[8],
            input_cells=["C6:D7"], linked_cells=["C10:D10"])
```
- rows의 **1-based 인덱스**로 섹션/소계/합계 지정, `currency_cols`=테마 통화단위, `input/linked_cells`=색 규약.
- 레이아웃: B2 제목·B4 헤더·B5~ 데이터(A열 여백). 반환=data_range.
- 세밀 제어가 필요하면 아래 개별 헬퍼(`apply_theme`·`style_total_row` 등)를 직접 호출.

### B-1. 표 스타일 (`excel_table_style`)
- **rules** : 격자 없음 + 헤더밑/소계/합계 가로선만. 헤더=흰바탕+primary 글씨.
- **grid**  : 외곽 medium + 내부 thin 전체 격자. 헤더=primary 채움+흰 글씨.
- **statement** : 격자 없음 + 표 **맨 위/맨 아래 medium 굵은선**. 헤더=회색 채움+검정 bold(상단 medium·하단 thin).

### B-2. 숫자 / 통화 서식 (`number_format_cols`)
음수=빨강 괄호(또는 슬레이트), 0=대시. 키 → 포맷:
| 키 | 포맷 | 예시 |
|---|---|---|
| `accounting`(권장) | `#,###,##0;[Red](#,###,##0);-` | 1,450,000 / (150,000) / - |
| `percent_acct` | `0.00%;[Red](0.00%);"-"` | 20.83% / (5.00%) / - |
| `thousands` | `#,##0;[Red](#,##0)` | 1,450,000 |
| `currency_won` | `#,##0"원";…` | 1,450,000원 |
| `million_won` | `#,##0,,"백만원"` | 1,450백만원 |
| `decimal` / `date` | `#,##0.00` / `yyyy-mm-dd` | 1,450.00 / 2026-06-02 |
| `multiple` | `0.0x` | 8.5x |
| `million` / `billion` | `#,##0,,` / `#,##0,,,` | 1,450 (백만/십억 스케일) |
| `change` / `bp` | `+#,##0;-#,##0;-` / `#,##0" bp"` | +250 / 125 bp |

> `write_table`의 `currency_cols`는 **테마 기본 통화서식**(전 테마 `accounting` 원 단위 회계서식, 2026-09-14)을 자동 적용. 백만원은 `number_cols` 로 `million_won` 지정.

### B-3. 테두리·정렬·행높이
- 테두리: 표 스타일에 따름(위 B-1). 정렬: 헤더=가운데(자동 줄바꿈), 본문=숫자 우측·텍스트 좌측·세로 가운데.
- 행높이: 테마별 `excel_row_header`/`excel_row_body` (예: dcf-valuation 22/18, default·procpa 20/16).

### B-4. 헤더·틀고정·줄무늬
- 헤더: 위 표 스타일. **틀 고정은 기본 OFF**(기본 규칙). 줄무늬: `excel_zebra` 테마값(공식 3종 모두 OFF).

### B-5. 소계·합계·섹션·하드코딩 (헬퍼)
- 소계/합계: `style_total_row(ws, theme, row, min_col, max_col, fill=)` — 스타일별 상/하선 + 굵게(+배경).
- 섹션 구분: `style_subheader_row(...)` — `subheader` 배경 + 굵게.
- **셀 종류 색 규약(DCF·감사)**: `mark_cells(ws, theme, input=["C5:D9"], linked=["C7"])` —
  입력(하드코딩)=크림(#FFFBEF) / 타시트 연결=파랑(#CCECFF) / 수식=검정. (`highlight_hardcoded`=input 별칭)

### B-6. 기본 규칙 (apply_theme 자동)
- **틀고정 안 함** / **A열·1행 여백**(너비 2.0·높이 ~27px) / **제목 색을 데이터 끝열까지 확장**(`title_cell`).
- **제목 밴드에는 시트명만**(default 계열, 확정 2026-09-10) — `write_title` 을 title 없이 부르면 시트명이 들어간다.
  설명·부제는 제목이 아니라 캡션·본문으로. audit 은 조서 헤더 6항목 그대로.
- 열너비 자동(한글 폭 2, **상한 `MAX_COL_WIDTH`=30** — 긴 문장은 wrap·병합), 음수 조건부 강조(`highlight_negatives`, 테마 `excel_highlight_neg`).

---


### B-7. 시트 탭 색 (`set_tab`, 사용자 확정 2026-09-10)
셀 색이 **셀 하나**의 성격(입력=크림·연결=연회색·수식=흰색)을 말한다면, 탭 색은 **시트 전체**의 성격을 시트를 열기 전에 말한다.
두 축은 서로 겹쳐도 모순되지 않는다 — 입력 시트의 탭은 회색 램프의 밝은 단, 그 안의 셀은 크림이다.

| 역할 | 이름 | 채우는 사람 | default·audit | procpa | dcf-valuation |
|---|---|---|---|---|---|
| `guide` | 안내·표지 | — | 색 없음 | 색 없음 | 색 없음 |
| `output` | 결과·총괄 | 자동(수식) | `393939` | `0B1C4A` | `004889` |
| `calc` | 계산·집계 | 자동(수식) | `808080` | `2563EB` | `2B579A` |
| `input` | 입력 — 작성자 | 작성자 | `A6A6A6` | `94A3B8` | `5B9BD5` |
| `pbc` | 입력 — 회사 제공 | 회사(PBC) | `BFBFBF` | `C9D4E5` | `D6E4F0` |
| `raw` | 원본·참고 | — | `D9D9D9` | `D9D9D9` | `D9D9D9` |

- **진하기 = 자리.** 진할수록 파일의 중심(결과), 연할수록 바깥(원천 자료). legacy(navy·charcoal)는 default 회색 램프를 쓴다.
- **무채색 = 우리 산출물이 아님.** `raw` 는 어느 테마에서든 `D9D9D9`. 받은 파일이므로 테마색을 주지 않는다.
- **새 색을 만들지 않는다.** 색 있는 테마는 팔레트가 이미 가진 토큰(title_bg·primary·secondary·secondary_lt·border_in)에서만 고른다.
  회색 테마에서는 그 램프가 마침 Office 표준 회색 5단과 정확히 일치한다.
- **안내·표지는 색 없음.** 색 없는 탭이 하나 있어야 나머지 색이 위계로 읽힌다.
- 색을 바꾸려면 `excel_theme.py` 의 `TAB_COLORS`(기본 램프)와 각 팔레트의 `"tabs"`(테마 오버라이드)만 고친다.

## C. 디자인 원칙 (생성 품질 — anti-slop)
- **절제된 단일 악센트**(한두 곳만) · **헤어라인 규율**(굵은 격자 대신 얇은 가로선) · **여백 > 장식** · **위계 우선**(제목>소계>본문).
- **색 60·30·10**: 60% 바탕/중립, 30% 보조, 10% 악센트. 의미 토큰만 쓰고 임의 hex 금지.
- **하지 말 것**: ❌ 과한 그라데이션·그림자 ❌ 이모지 불릿(• OK) ❌ 의미 없는 색칠/무지개 표
  ❌ 원본에 없는 장식선(dcf-valuation에 오렌지선 추가 금지) ❌ 빽빽한 전체격자+짙은 배경 동시.
- **타이포**: 제목/헤딩만 크게·굵게, 본문은 한 색(text). 폰트 폴백 Pretendard→맑은 고딕.
- **접근성**: 헤더/합계 글자 vs 배경 **대비** 확보, 색만으로 의미 구분 금지(부호·괄호 병행).
- **앱별**: 엑셀=여백·틀고정 끔·제목바 확장 / 워드=H1>H2>H3·콜아웃 1개·페이지번호 / PPT=한 슬라이드 한 메시지·장식선 금지.

---

## D. 수정·미리보기 워크플로
1. **색/폰트/표규칙 변경** — `excel_theme.py` 상단 `PALETTES`에서 hex·`font_*`·`sizes`·`excel_*` 수정.
   (또는 엑셀로 편집: `palette_io.py` = `dev.ps1 -Export`/`-Import`)
2. **빌드** — `python build_all.py` → `palette_swatch.html`(색 일람) + 각 앱 `examples/preview_*` 생성.
3. **확인** — 미리보기를 열어 보고, 마음에 안 들면 1로 반복.
4. **새 테마 추가** — `PALETTES`에 항목 추가(의미 토큰 primary/text/band/accent…). 기본 테마는 `DEFAULT_THEME` 한 줄.
   저장 후 `apply_theme(ws, theme="새이름")` 으로 바로 사용(워드/PPT도 자동 반영).

> 참고: 테마 발전 도구는 `tools/` (sample_gen → 엑셀에서 직접 수정 → style_diff 자동 반영 / analyze_reference 실무파일 분석). 프로젝트 폴더 DEV.md 의 루프 설명 참조.
